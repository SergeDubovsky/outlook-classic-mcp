using System;
using System.Security;
using System.Security.Cryptography;
using System.Text;

namespace OutlookClassicMcp.Transport
{
    public sealed class BearerToken : IDisposable
    {
        public const string EnvironmentVariableName = "OUTLOOK_MCP_TOKEN";
        public const int EncodedLength = 43;
        private const int DecodedLength = 32;
        private const string CursorKeyDomain = "outlook-classic-mcp/cursor-key/v1";
        private readonly object _gate = new object();
        private byte[]? _value;

        private BearerToken(byte[] value)
        {
            _value = value;
        }

        public static BearerToken LoadFromProcessEnvironment()
        {
            var configured = Environment.GetEnvironmentVariable(
                EnvironmentVariableName,
                EnvironmentVariableTarget.Process);
            if (!TryCreate(configured, out var token))
            {
                throw new SecurityException(
                    $"{EnvironmentVariableName} must contain one canonical 32-byte base64url token. Restart Outlook after provisioning or rotating it.");
            }

            return token;
        }

        public static bool TryCreate(string? encoded, out BearerToken token)
        {
            token = null!;
            if (!TryDecode(encoded, out var value))
            {
                return false;
            }

            token = new BearerToken(value);
            return true;
        }

        public bool MatchesAuthorizationHeaders(string[]? values)
        {
            lock (_gate)
            {
                var expected = _value ?? throw new ObjectDisposedException(nameof(BearerToken));
                if (values == null || values.Length != 1)
                {
                    return false;
                }

                const string scheme = "Bearer ";
                var header = values[0];
                if (header == null ||
                    header.Length != scheme.Length + EncodedLength ||
                    !header.StartsWith(scheme, StringComparison.OrdinalIgnoreCase) ||
                    !TryDecode(header.Substring(scheme.Length), out var candidate))
                {
                    return false;
                }

                try
                {
                    var difference = 0;
                    for (var index = 0; index < DecodedLength; index++)
                    {
                        difference |= expected[index] ^ candidate[index];
                    }

                    return difference == 0;
                }
                finally
                {
                    Array.Clear(candidate, 0, candidate.Length);
                }
            }
        }

        internal HmacCursorCodec CreateCursorCodec()
        {
            byte[] domainBytes = Array.Empty<byte>();
            byte[] derivedKey = Array.Empty<byte>();
            try
            {
                lock (_gate)
                {
                    var value = _value ?? throw new ObjectDisposedException(nameof(BearerToken));
                    domainBytes = Encoding.UTF8.GetBytes(CursorKeyDomain);
                    using (var hmac = new HMACSHA256(value))
                    {
                        derivedKey = hmac.ComputeHash(domainBytes);
                    }
                }

                return new HmacCursorCodec(derivedKey);
            }
            finally
            {
                Array.Clear(domainBytes, 0, domainBytes.Length);
                Array.Clear(derivedKey, 0, derivedKey.Length);
            }
        }

        public void Dispose()
        {
            byte[]? value;
            lock (_gate)
            {
                value = _value;
                _value = null;
            }

            if (value != null)
            {
                Array.Clear(value, 0, value.Length);
            }
        }

        private static bool TryDecode(string? encoded, out byte[] value)
        {
            value = Array.Empty<byte>();
            if (encoded == null || encoded.Length != EncodedLength)
            {
                return false;
            }

            for (var index = 0; index < encoded.Length; index++)
            {
                var character = encoded[index];
                if (!((character >= 'A' && character <= 'Z') ||
                      (character >= 'a' && character <= 'z') ||
                      (character >= '0' && character <= '9') ||
                      character == '-' || character == '_'))
                {
                    return false;
                }
            }

            try
            {
                var base64 = encoded.Replace('-', '+').Replace('_', '/') + "=";
                var decoded = Convert.FromBase64String(base64);
                if (decoded.Length != DecodedLength ||
                    !string.Equals(Encode(decoded), encoded, StringComparison.Ordinal))
                {
                    Array.Clear(decoded, 0, decoded.Length);
                    return false;
                }

                value = decoded;
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }

        private static string Encode(byte[] value)
        {
            return Convert.ToBase64String(value)
                .TrimEnd('=')
                .Replace('+', '-')
                .Replace('/', '_');
        }
    }
}
