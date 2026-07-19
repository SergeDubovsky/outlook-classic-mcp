using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using NUnit.Framework;

namespace OutlookClassicMcp.Transport.Tests
{
    [TestFixture]
    public sealed class HttpRequestValidatorTests
    {
        private const string TokenText = "AAECAwQFBgcICQoLDA0ODxAREhMUFRYXGBkaGxwdHh8";

        [Test]
        public void AcceptsCanonicalRequest()
        {
            var decision = Validate(CreateFacts());

            Assert.That(decision.IsAccepted, Is.True);
            Assert.That(decision.StatusCode, Is.Zero);
            Assert.That(decision.ErrorCategory, Is.EqualTo(HttpRequestErrorCategory.None));
            Assert.That(decision.RequiredHeaders, Is.Empty);
        }

        [TestCase("192.168.1.10")]
        [TestCase("127.0.0.2")]
        [TestCase("::1")]
        [TestCase("::ffff:127.0.0.1")]
        public void RequiresExactIpv4LoopbackRemoteAddress(string address)
        {
            var facts = CreateFacts(remoteAddress: IPAddress.Parse(address));

            AssertRejected(
                Validate(facts),
                403,
                HttpRequestErrorCategory.NonLoopbackRemoteEndpoint);
        }

        [Test]
        public void RejectsMissingRemoteAddress()
        {
            var facts = new HttpRequestFacts(null, "/mcp/", "POST", 0, CreateDefaultHeaders());

            AssertRejected(
                Validate(facts),
                403,
                HttpRequestErrorCategory.NonLoopbackRemoteEndpoint);
        }

        [TestCase("localhost:8765")]
        [TestCase("127.0.0.1")]
        [TestCase("127.0.0.1:80")]
        [TestCase("127.0.0.1:8765 ")]
        public void RequiresExactHost(string host)
        {
            var facts = CreateFacts(headers: ReplaceHeader("Host", new HttpHeaderFact("Host", host)));

            AssertRejected(Validate(facts), 400, HttpRequestErrorCategory.InvalidHost);
        }

        [Test]
        public void RejectsMissingOrDuplicateHost()
        {
            var missing = CreateFacts(headers: WithoutHeader("Host"));
            var duplicateHeaders = CreateDefaultHeaders().ToList();
            duplicateHeaders.Add(new HttpHeaderFact("host", HttpRequestValidator.CanonicalHost));

            AssertRejected(Validate(missing), 400, HttpRequestErrorCategory.InvalidHost);
            AssertRejected(
                Validate(CreateFacts(headers: duplicateHeaders)),
                400,
                HttpRequestErrorCategory.InvalidHost);
        }

        [TestCase("/")]
        [TestCase("/mcp")]
        [TestCase("/mcp//")]
        [TestCase("/mcp/tools")]
        [TestCase("/mcp/?x=1")]
        [TestCase("/%6dcp/")]
        [TestCase("/mcp/../mcp/")]
        public void RequiresExactRawRouteWithoutQueryOrNormalization(string rawUrl)
        {
            AssertRejected(
                Validate(CreateFacts(rawUrl: rawUrl)),
                404,
                HttpRequestErrorCategory.InvalidRoute);
        }

        [TestCase("GET")]
        [TestCase("DELETE")]
        [TestCase("HEAD")]
        [TestCase("OPTIONS")]
        [TestCase("post")]
        public void RejectsEveryNonCanonicalPostMethod(string method)
        {
            var decision = Validate(CreateFacts(method: method));

            AssertRejected(decision, 405, HttpRequestErrorCategory.MethodNotAllowed);
            Assert.That(decision.RequiredHeaders, Contains.Key("Allow"));
            Assert.That(decision.RequiredHeaders["Allow"], Is.EqualTo("POST"));
        }

        [Test]
        public void RejectsOriginByPresenceIncludingEmptyValue()
        {
            var headers = CreateDefaultHeaders().ToList();
            headers.Add(new HttpHeaderFact("Origin", string.Empty));

            AssertRejected(
                Validate(CreateFacts(headers: headers)),
                403,
                HttpRequestErrorCategory.OriginRejected);
        }

        [Test]
        public void EveryAuthorizationFailureHasTheSameDecision()
        {
            IEnumerable<HttpHeaderFact>[] cases =
            {
                WithoutHeader("Authorization"),
                ReplaceHeader("Authorization", new HttpHeaderFact("Authorization", "Basic " + TokenText)),
                ReplaceHeader("Authorization", new HttpHeaderFact("Authorization", "Bearer " + new string('A', 43))),
                ReplaceHeader("Authorization", new HttpHeaderFact("Authorization", "Bearer  " + TokenText)),
                ReplaceHeader(
                    "Authorization",
                    new HttpHeaderFact("Authorization", "Bearer " + TokenText, "Bearer " + TokenText)),
            };

            foreach (var headers in cases)
            {
                var decision = Validate(CreateFacts(headers: headers));
                AssertRejected(decision, 401, HttpRequestErrorCategory.Unauthorized);
                Assert.That(decision.RequiredHeaders.Count, Is.EqualTo(1));
                Assert.That(decision.RequiredHeaders["WWW-Authenticate"], Is.EqualTo("Bearer"));
            }
        }

        [Test]
        public void EnforcesHeaderFieldValueAndAggregateBudgets()
        {
            var exactFieldCount = CreateDefaultHeaders().ToList();
            while (CountFields(exactFieldCount) < RequestLimits.MaximumHeaderFields)
            {
                exactFieldCount.Add(new HttpHeaderFact("X", "1"));
            }

            Assert.That(Validate(CreateFacts(headers: exactFieldCount)).IsAccepted, Is.True);

            var tooManyFields = exactFieldCount.ToList();
            tooManyFields.Add(new HttpHeaderFact("X", "1"));
            AssertRejected(
                Validate(CreateFacts(headers: tooManyFields)),
                431,
                HttpRequestErrorCategory.HeadersTooLarge);

            var maximumValue = CreateDefaultHeaders().ToList();
            maximumValue.Add(
                new HttpHeaderFact("X-Max", new string('x', RequestLimits.MaximumHeaderValueCharacters)));
            Assert.That(Validate(CreateFacts(headers: maximumValue)).IsAccepted, Is.True);

            var oversizedValue = CreateDefaultHeaders().ToList();
            oversizedValue.Add(
                new HttpHeaderFact("X-Max", new string('x', RequestLimits.MaximumHeaderValueCharacters + 1)));
            AssertRejected(
                Validate(CreateFacts(headers: oversizedValue)),
                431,
                HttpRequestErrorCategory.HeadersTooLarge);

            var oversizedAggregate = CreateDefaultHeaders().ToList();
            oversizedAggregate.Add(new HttpHeaderFact("X-One", new string('x', 8100)));
            oversizedAggregate.Add(new HttpHeaderFact("X-Two", new string('x', 8101)));
            AssertRejected(
                Validate(CreateFacts(headers: oversizedAggregate)),
                431,
                HttpRequestErrorCategory.HeadersTooLarge);
        }

        [TestCase("application/json")]
        [TestCase("Application/Json; Charset=UTF-8")]
        [TestCase("application/json; charset=\"utf-8\"; profile=example")]
        public void AcceptsApplicationJsonWithSafeOptionalParameters(string contentType)
        {
            var facts = CreateFacts(
                headers: ReplaceHeader("Content-Type", new HttpHeaderFact("Content-Type", contentType)));

            Assert.That(Validate(facts).IsAccepted, Is.True);
        }

        [TestCase("text/plain")]
        [TestCase("application/json; charset=iso-8859-1")]
        [TestCase("application/json; charset=utf-8; charset=utf-8")]
        [TestCase("application/json,")]
        public void RejectsUnsupportedOrAmbiguousContentType(string contentType)
        {
            AssertRejected(
                Validate(
                    CreateFacts(
                        headers: ReplaceHeader(
                            "Content-Type",
                            new HttpHeaderFact("Content-Type", contentType)))),
                415,
                HttpRequestErrorCategory.UnsupportedContentType);
        }

        [Test]
        public void RejectsMissingOrDuplicateContentType()
        {
            AssertRejected(
                Validate(CreateFacts(headers: WithoutHeader("Content-Type"))),
                415,
                HttpRequestErrorCategory.UnsupportedContentType);
            AssertRejected(
                Validate(
                    CreateFacts(
                        headers: ReplaceHeader(
                            "Content-Type",
                            new HttpHeaderFact("Content-Type", "application/json", "application/json")))),
                415,
                HttpRequestErrorCategory.UnsupportedContentType);
        }

        [TestCase("application/json, text/event-stream")]
        [TestCase("TEXT/EVENT-STREAM; q=0.4, APPLICATION/JSON; q=1")]
        [TestCase("application/json; profile=\"a,b\", text/event-stream")]
        public void AcceptsRequiredMediaTypesIndependentOfOrderCaseAndParameters(string accept)
        {
            var facts = CreateFacts(headers: ReplaceHeader("Accept", new HttpHeaderFact("Accept", accept)));

            Assert.That(Validate(facts).IsAccepted, Is.True);
        }

        [Test]
        public void AcceptsRequiredMediaTypesAcrossHeaderValues()
        {
            var facts = CreateFacts(
                headers: ReplaceHeader(
                    "Accept",
                    new HttpHeaderFact("Accept", "application/json", "text/event-stream")));

            Assert.That(Validate(facts).IsAccepted, Is.True);
        }

        [TestCase("application/json")]
        [TestCase("text/event-stream")]
        [TestCase("*/*")]
        [TestCase("application/json, text/event-stream; q=0")]
        [TestCase("application/json, text/event-stream; q=bogus")]
        [TestCase("application/json, text/event-stream,")]
        public void RejectsIncompleteWildcardZeroQualityOrMalformedAccept(string accept)
        {
            AssertRejected(
                Validate(
                    CreateFacts(
                        headers: ReplaceHeader("Accept", new HttpHeaderFact("Accept", accept)))),
                406,
                HttpRequestErrorCategory.NotAcceptable);
        }

        [TestCase("2024-11-05")]
        [TestCase("2025-03-26")]
        [TestCase("2025-06-18")]
        [TestCase("2025-11-25")]
        public void AcceptsPinnedSdkProtocolVersions(string protocolVersion)
        {
            var facts = CreateFacts(
                headers: AddHeader(
                    CreateDefaultHeaders(),
                    new HttpHeaderFact("MCP-Protocol-Version", protocolVersion)));

            Assert.That(Validate(facts).IsAccepted, Is.True);
        }

        [Test]
        public void AcceptsMissingProtocolVersionForNegotiationOrFallback()
        {
            Assert.That(Validate(CreateFacts()).IsAccepted, Is.True);
        }

        [Test]
        public void RejectsUnsupportedEmptyOrDuplicateProtocolVersion()
        {
            foreach (var header in new[]
            {
                new HttpHeaderFact("MCP-Protocol-Version", "2020-01-01"),
                new HttpHeaderFact("MCP-Protocol-Version", string.Empty),
                new HttpHeaderFact("MCP-Protocol-Version", "2025-03-26", "2025-03-26"),
            })
            {
                AssertRejected(
                    Validate(CreateFacts(headers: AddHeader(CreateDefaultHeaders(), header))),
                    400,
                    HttpRequestErrorCategory.UnsupportedProtocolVersion);
            }
        }

        [TestCase("identity")]
        [TestCase("IDENTITY")]
        public void AcceptsExplicitIdentityContentEncoding(string encoding)
        {
            var facts = CreateFacts(
                headers: AddHeader(CreateDefaultHeaders(), new HttpHeaderFact("Content-Encoding", encoding)));

            Assert.That(Validate(facts).IsAccepted, Is.True);
        }

        [Test]
        public void RejectsCompressedOrDuplicateContentEncoding()
        {
            foreach (var header in new[]
            {
                new HttpHeaderFact("Content-Encoding", "gzip"),
                new HttpHeaderFact("Content-Encoding", "identity", "identity"),
            })
            {
                AssertRejected(
                    Validate(CreateFacts(headers: AddHeader(CreateDefaultHeaders(), header))),
                    415,
                    HttpRequestErrorCategory.UnsupportedContentEncoding);
            }
        }

        [Test]
        public void EnforcesDeclaredBodyLimitInBytes()
        {
            Assert.That(
                Validate(CreateFacts(contentLength: RequestLimits.MaximumRequestBodyBytes)).IsAccepted,
                Is.True);
            Assert.That(Validate(CreateFacts(contentLength: -1)).IsAccepted, Is.True);
            AssertRejected(
                Validate(CreateFacts(contentLength: RequestLimits.MaximumRequestBodyBytes + 1)),
                413,
                HttpRequestErrorCategory.PayloadTooLarge);
            AssertRejected(
                Validate(CreateFacts(contentLength: -2)),
                400,
                HttpRequestErrorCategory.InvalidContentLength);
        }

        [Test]
        public void RejectsAnySessionHeaderInStatelessMode()
        {
            foreach (var header in new[]
            {
                new HttpHeaderFact("Mcp-Session-Id", string.Empty),
                new HttpHeaderFact("Mcp-Session-Id", "session"),
            })
            {
                AssertRejected(
                    Validate(CreateFacts(headers: AddHeader(CreateDefaultHeaders(), header))),
                    400,
                    HttpRequestErrorCategory.SessionHeaderNotAllowed);
            }
        }

        [Test]
        public void StopsAtTheFirstFailureInTheDocumentedOrder()
        {
            var headers = new List<HttpHeaderFact>
            {
                new HttpHeaderFact("Host", "localhost:8765"),
                new HttpHeaderFact("Origin", "https://example.invalid"),
                new HttpHeaderFact("Authorization", "Basic invalid"),
                new HttpHeaderFact("Content-Type", "text/plain"),
                new HttpHeaderFact("Accept", "*/*"),
                new HttpHeaderFact("MCP-Protocol-Version", "invalid"),
                new HttpHeaderFact("Content-Encoding", "gzip"),
                new HttpHeaderFact("Mcp-Session-Id", "session"),
            };
            var facts = new HttpRequestFacts(
                IPAddress.Parse("192.168.1.10"),
                "/wrong",
                "GET",
                RequestLimits.MaximumRequestBodyBytes + 1,
                headers);

            AssertRejected(
                Validate(facts),
                403,
                HttpRequestErrorCategory.NonLoopbackRemoteEndpoint);

            var validRemote = new HttpRequestFacts(
                IPAddress.Loopback,
                "/wrong",
                "GET",
                RequestLimits.MaximumRequestBodyBytes + 1,
                headers);
            AssertRejected(Validate(validRemote), 400, HttpRequestErrorCategory.InvalidHost);
        }

        [Test]
        public void AuthorizationPrecedesHeaderBudgetAndBodyLimitPrecedesSessionHeader()
        {
            var unauthorizedHeaders = ReplaceHeader(
                "Authorization",
                new HttpHeaderFact("Authorization", "Basic invalid")).ToList();
            unauthorizedHeaders.Add(
                new HttpHeaderFact("X-Large", new string('x', RequestLimits.MaximumHeaderValueCharacters + 1)));
            AssertRejected(
                Validate(CreateFacts(headers: unauthorizedHeaders)),
                401,
                HttpRequestErrorCategory.Unauthorized);

            var sessionHeaders = AddHeader(
                CreateDefaultHeaders(),
                new HttpHeaderFact("Mcp-Session-Id", "session"));
            AssertRejected(
                Validate(
                    CreateFacts(
                        headers: sessionHeaders,
                        contentLength: RequestLimits.MaximumRequestBodyBytes + 1)),
                413,
                HttpRequestErrorCategory.PayloadTooLarge);
        }

        private static HttpRequestValidationDecision Validate(HttpRequestFacts facts)
        {
            Assert.That(BearerToken.TryCreate(TokenText, out var token), Is.True);
            using (token)
            {
                return new HttpRequestValidator(token).Validate(facts);
            }
        }

        private static HttpRequestFacts CreateFacts(
            IEnumerable<HttpHeaderFact>? headers = null,
            IPAddress? remoteAddress = null,
            string rawUrl = "/mcp/",
            string method = "POST",
            long contentLength = 0)
        {
            return new HttpRequestFacts(
                remoteAddress ?? IPAddress.Loopback,
                rawUrl,
                method,
                contentLength,
                headers ?? CreateDefaultHeaders());
        }

        private static HttpHeaderFact[] CreateDefaultHeaders()
        {
            return new[]
            {
                new HttpHeaderFact("Host", HttpRequestValidator.CanonicalHost),
                new HttpHeaderFact("Authorization", "Bearer " + TokenText),
                new HttpHeaderFact("Content-Type", "application/json"),
                new HttpHeaderFact("Accept", "application/json, text/event-stream"),
            };
        }

        private static List<HttpHeaderFact> ReplaceHeader(
            string name,
            HttpHeaderFact replacement)
        {
            var headers = WithoutHeader(name).ToList();
            headers.Add(replacement);
            return headers;
        }

        private static HttpHeaderFact[] WithoutHeader(string name)
        {
            return CreateDefaultHeaders()
                .Where(header => !string.Equals(header.Name, name, StringComparison.OrdinalIgnoreCase))
                .ToArray();
        }

        private static HttpHeaderFact[] AddHeader(
            IEnumerable<HttpHeaderFact> headers,
            HttpHeaderFact header)
        {
            return headers.Concat(new[] { header }).ToArray();
        }

        private static int CountFields(IEnumerable<HttpHeaderFact> headers)
        {
            return headers.Sum(header => Math.Max(1, header.Values.Count));
        }

        private static void AssertRejected(
            HttpRequestValidationDecision decision,
            int statusCode,
            HttpRequestErrorCategory category)
        {
            Assert.That(decision.IsAccepted, Is.False);
            Assert.That(decision.StatusCode, Is.EqualTo(statusCode));
            Assert.That(decision.ErrorCategory, Is.EqualTo(category));
        }
    }
}
