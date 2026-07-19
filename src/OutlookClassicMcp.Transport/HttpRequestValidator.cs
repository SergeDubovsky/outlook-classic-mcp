using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net;
using System.Net.Http.Headers;

namespace OutlookClassicMcp.Transport
{
    public enum HttpRequestErrorCategory
    {
        None = 0,
        NonLoopbackRemoteEndpoint,
        InvalidHost,
        InvalidRoute,
        MethodNotAllowed,
        OriginRejected,
        Unauthorized,
        HeadersTooLarge,
        UnsupportedContentType,
        NotAcceptable,
        UnsupportedProtocolVersion,
        UnsupportedContentEncoding,
        InvalidContentLength,
        PayloadTooLarge,
        SessionHeaderNotAllowed,
    }

    public sealed class HttpRequestValidationDecision
    {
        private static readonly IReadOnlyDictionary<string, string> NoHeaders =
            new ReadOnlyDictionary<string, string>(
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase));

        internal HttpRequestValidationDecision(
            bool isAccepted,
            int statusCode,
            HttpRequestErrorCategory errorCategory,
            IReadOnlyDictionary<string, string>? requiredHeaders = null)
        {
            IsAccepted = isAccepted;
            StatusCode = statusCode;
            ErrorCategory = errorCategory;
            RequiredHeaders = requiredHeaders ?? NoHeaders;
        }

        public bool IsAccepted { get; }

        public int StatusCode { get; }

        public HttpRequestErrorCategory ErrorCategory { get; }

        public IReadOnlyDictionary<string, string> RequiredHeaders { get; }
    }

    public sealed class HttpRequestValidator
    {
        public const string CanonicalHost = "127.0.0.1:8765";
        public const string CanonicalRoute = "/mcp/";

        private const string ApplicationJson = "application/json";
        private const string EventStream = "text/event-stream";
        private static readonly HashSet<string> SupportedProtocolVersions =
            new HashSet<string>(StringComparer.Ordinal)
            {
                "2024-11-05",
                "2025-03-26",
                "2025-06-18",
                "2025-11-25",
            };

        private readonly BearerToken _bearerToken;

        public HttpRequestValidator(BearerToken bearerToken)
        {
            _bearerToken = bearerToken ?? throw new ArgumentNullException(nameof(bearerToken));
        }

        public HttpRequestValidationDecision Validate(HttpRequestFacts facts)
        {
            if (facts == null)
            {
                throw new ArgumentNullException(nameof(facts));
            }

            if (!IPAddress.Loopback.Equals(facts.RemoteAddress))
            {
                return Reject(HttpStatusCode.Forbidden, HttpRequestErrorCategory.NonLoopbackRemoteEndpoint);
            }

            var hostValues = facts.GetHeaderValues("Host");
            if (hostValues == null ||
                hostValues.Length != 1 ||
                !string.Equals(hostValues[0], CanonicalHost, StringComparison.Ordinal))
            {
                return Reject(HttpStatusCode.BadRequest, HttpRequestErrorCategory.InvalidHost);
            }

            if (!string.Equals(facts.RawUrl, CanonicalRoute, StringComparison.Ordinal))
            {
                return Reject(HttpStatusCode.NotFound, HttpRequestErrorCategory.InvalidRoute);
            }

            if (!string.Equals(facts.Method, "POST", StringComparison.Ordinal))
            {
                return Reject(
                    HttpStatusCode.MethodNotAllowed,
                    HttpRequestErrorCategory.MethodNotAllowed,
                    "Allow",
                    "POST");
            }

            if (facts.HasHeader("Origin"))
            {
                return Reject(HttpStatusCode.Forbidden, HttpRequestErrorCategory.OriginRejected);
            }

            if (!_bearerToken.MatchesAuthorizationHeaders(facts.GetHeaderValues("Authorization")))
            {
                return Reject(
                    HttpStatusCode.Unauthorized,
                    HttpRequestErrorCategory.Unauthorized,
                    "WWW-Authenticate",
                    "Bearer");
            }

            if (!AreHeadersWithinLimits(facts.Headers))
            {
                return Reject(431, HttpRequestErrorCategory.HeadersTooLarge);
            }

            if (!HasSupportedContentType(facts.GetHeaderValues("Content-Type")))
            {
                return Reject(HttpStatusCode.UnsupportedMediaType, HttpRequestErrorCategory.UnsupportedContentType);
            }

            if (!HasRequiredAcceptMediaTypes(facts.GetHeaderValues("Accept")))
            {
                return Reject(HttpStatusCode.NotAcceptable, HttpRequestErrorCategory.NotAcceptable);
            }

            if (!HasSupportedProtocolVersion(facts.GetHeaderValues("MCP-Protocol-Version")))
            {
                return Reject(HttpStatusCode.BadRequest, HttpRequestErrorCategory.UnsupportedProtocolVersion);
            }

            if (!HasSupportedContentEncoding(facts.GetHeaderValues("Content-Encoding")))
            {
                return Reject(HttpStatusCode.UnsupportedMediaType, HttpRequestErrorCategory.UnsupportedContentEncoding);
            }

            if (facts.ContentLength < -1)
            {
                return Reject(HttpStatusCode.BadRequest, HttpRequestErrorCategory.InvalidContentLength);
            }

            if (facts.ContentLength > RequestLimits.MaximumRequestBodyBytes)
            {
                return Reject(HttpStatusCode.RequestEntityTooLarge, HttpRequestErrorCategory.PayloadTooLarge);
            }

            if (facts.HasHeader("Mcp-Session-Id"))
            {
                return Reject(HttpStatusCode.BadRequest, HttpRequestErrorCategory.SessionHeaderNotAllowed);
            }

            return new HttpRequestValidationDecision(true, 0, HttpRequestErrorCategory.None);
        }

        private static bool AreHeadersWithinLimits(IReadOnlyList<HttpHeaderFact> headers)
        {
            var fieldCount = 0;
            long characterCount = 0;
            foreach (var header in headers)
            {
                if (header.Values.Count == 0)
                {
                    fieldCount++;
                    characterCount += header.Name.Length + 4L;
                }
                else
                {
                    foreach (var value in header.Values)
                    {
                        fieldCount++;
                        if (value.Length > RequestLimits.MaximumHeaderValueCharacters)
                        {
                            return false;
                        }

                        characterCount += header.Name.Length + value.Length + 4L;
                    }
                }

                if (fieldCount > RequestLimits.MaximumHeaderFields ||
                    characterCount > RequestLimits.MaximumHeaderCharacters)
                {
                    return false;
                }
            }

            return true;
        }

        private static bool HasSupportedContentType(string[]? values)
        {
            if (values == null ||
                values.Length != 1 ||
                !MediaTypeHeaderValue.TryParse(values[0], out var contentType) ||
                !string.Equals(contentType.MediaType, ApplicationJson, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            var charsets = contentType.Parameters
                .Where(parameter => string.Equals(parameter.Name, "charset", StringComparison.OrdinalIgnoreCase))
                .Select(parameter => parameter.Value?.Trim('"'))
                .ToArray();
            return charsets.Length <= 1 &&
                (charsets.Length == 0 || string.Equals(charsets[0], "utf-8", StringComparison.OrdinalIgnoreCase));
        }

        private static bool HasRequiredAcceptMediaTypes(string[]? values)
        {
            if (values == null || values.Length == 0)
            {
                return false;
            }

            var hasApplicationJson = false;
            var hasEventStream = false;
            foreach (var value in SplitHeaderValues(values))
            {
                if (!MediaTypeWithQualityHeaderValue.TryParse(value, out var mediaType))
                {
                    return false;
                }

                var qualityParameters = mediaType.Parameters
                    .Where(parameter => string.Equals(parameter.Name, "q", StringComparison.OrdinalIgnoreCase))
                    .ToArray();
                if (qualityParameters.Length > 1)
                {
                    return false;
                }

                double? quality;
                try
                {
                    quality = mediaType.Quality;
                }
                catch (FormatException)
                {
                    return false;
                }

                if (qualityParameters.Length == 1 && !quality.HasValue)
                {
                    return false;
                }

                if (quality.HasValue && quality.Value <= 0)
                {
                    continue;
                }

                hasApplicationJson |= string.Equals(
                    mediaType.MediaType,
                    ApplicationJson,
                    StringComparison.OrdinalIgnoreCase);
                hasEventStream |= string.Equals(
                    mediaType.MediaType,
                    EventStream,
                    StringComparison.OrdinalIgnoreCase);
            }

            return hasApplicationJson && hasEventStream;
        }

        private static bool HasSupportedProtocolVersion(string[]? values)
        {
            return values == null ||
                (values.Length == 1 && SupportedProtocolVersions.Contains(values[0]));
        }

        private static bool HasSupportedContentEncoding(string[]? values)
        {
            return values == null ||
                (values.Length == 1 && string.Equals(values[0], "identity", StringComparison.OrdinalIgnoreCase));
        }

        private static IEnumerable<string> SplitHeaderValues(IEnumerable<string> values)
        {
            foreach (var value in values)
            {
                var start = 0;
                var quoted = false;
                var escaped = false;
                for (var index = 0; index < value.Length; index++)
                {
                    var character = value[index];
                    if (escaped)
                    {
                        escaped = false;
                    }
                    else if (quoted && character == '\\')
                    {
                        escaped = true;
                    }
                    else if (character == '"')
                    {
                        quoted = !quoted;
                    }
                    else if (!quoted && character == ',')
                    {
                        yield return value.Substring(start, index - start).Trim();
                        start = index + 1;
                    }
                }

                if (quoted || escaped)
                {
                    yield return string.Empty;
                }
                else
                {
                    yield return value.Substring(start).Trim();
                }
            }
        }

        private static HttpRequestValidationDecision Reject(
            HttpStatusCode statusCode,
            HttpRequestErrorCategory category)
        {
            return Reject((int)statusCode, category);
        }

        private static HttpRequestValidationDecision Reject(
            int statusCode,
            HttpRequestErrorCategory category)
        {
            return new HttpRequestValidationDecision(false, statusCode, category);
        }

        private static HttpRequestValidationDecision Reject(
            HttpStatusCode statusCode,
            HttpRequestErrorCategory category,
            string headerName,
            string headerValue)
        {
            var headers = new ReadOnlyDictionary<string, string>(
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    [headerName] = headerValue,
                });
            return new HttpRequestValidationDecision(false, (int)statusCode, category, headers);
        }
    }
}
