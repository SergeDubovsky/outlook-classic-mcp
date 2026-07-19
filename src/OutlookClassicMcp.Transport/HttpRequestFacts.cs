using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net;

namespace OutlookClassicMcp.Transport
{
    public sealed class HttpHeaderFact
    {
        private readonly ReadOnlyCollection<string> _values;

        public HttpHeaderFact(string name, params string[] values)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new ArgumentException("A header name is required.", nameof(name));
            }

            if (values == null)
            {
                throw new ArgumentNullException(nameof(values));
            }

            if (values.Any(value => value == null))
            {
                throw new ArgumentException("Header values cannot contain null.", nameof(values));
            }

            Name = name;
            _values = Array.AsReadOnly((string[])values.Clone());
        }

        public string Name { get; }

        public IReadOnlyList<string> Values => _values;
    }

    public sealed class HttpRequestFacts
    {
        private readonly ReadOnlyCollection<HttpHeaderFact> _headers;

        public HttpRequestFacts(
            IPAddress? remoteAddress,
            string? rawUrl,
            string? method,
            long contentLength,
            IEnumerable<HttpHeaderFact> headers)
        {
            if (headers == null)
            {
                throw new ArgumentNullException(nameof(headers));
            }

            RemoteAddress = remoteAddress;
            RawUrl = rawUrl;
            Method = method;
            ContentLength = contentLength;
            _headers = Array.AsReadOnly(headers.ToArray());
        }

        public IPAddress? RemoteAddress { get; }

        public string? RawUrl { get; }

        public string? Method { get; }

        public long ContentLength { get; }

        public IReadOnlyList<HttpHeaderFact> Headers => _headers;

        public bool HasHeader(string name)
        {
            if (name == null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            return _headers.Any(header =>
                string.Equals(header.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        public string[]? GetHeaderValues(string name)
        {
            if (name == null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            var values = _headers
                .Where(header => string.Equals(header.Name, name, StringComparison.OrdinalIgnoreCase))
                .SelectMany(header => header.Values)
                .ToArray();
            return values.Length == 0 && !HasHeader(name) ? null : values;
        }
    }
}
