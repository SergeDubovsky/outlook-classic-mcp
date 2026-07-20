using System;
using System.Security;
using NUnit.Framework;

namespace OutlookClassicMcp.Transport.Tests
{
    [TestFixture]
    [NonParallelizable]
    public sealed class BearerTokenTests
    {
        private const string TokenText = "AAECAwQFBgcICQoLDA0ODxAREhMUFRYXGBkaGxwdHh8";

        [Test]
        public void AcceptsExactlyOneCanonicalBearerValue()
        {
            using (var token = CreateToken())
            {
                Assert.That(token.MatchesAuthorizationHeaders(new[] { "Bearer " + TokenText }), Is.True);
                Assert.That(token.MatchesAuthorizationHeaders(new[] { "bearer " + TokenText }), Is.True);
            }
        }

        [TestCase(null)]
        [TestCase("")]
        [TestCase("short")]
        [TestCase("AAECAwQFBgcICQoLDA0ODxAREhMUFRYXGBkaGxwdHh=")]
        [TestCase("AAECAwQFBgcICQoLDA0ODxAREhMUFRYXGBkaGxwdHh+")]
        public void RejectsNonCanonicalTokenConfiguration(string? configured)
        {
            Assert.That(BearerToken.TryCreate(configured, out _), Is.False);
        }

        [Test]
        public void FailsClosedForMissingWrongDuplicateAndMalformedAuthorization()
        {
            using (var token = CreateToken())
            {
                Assert.That(token.MatchesAuthorizationHeaders(null), Is.False);
                Assert.That(token.MatchesAuthorizationHeaders(Array.Empty<string>()), Is.False);
                Assert.That(token.MatchesAuthorizationHeaders(new[] { "Basic " + TokenText }), Is.False);
                Assert.That(token.MatchesAuthorizationHeaders(new[] { "Bearer  " + TokenText }), Is.False);
                Assert.That(token.MatchesAuthorizationHeaders(new[] { "Bearer " + TokenText + " " }), Is.False);
                Assert.That(token.MatchesAuthorizationHeaders(new[] { "Bearer " + new string('A', 43) }), Is.False);
                Assert.That(
                    token.MatchesAuthorizationHeaders(
                        new[] { "Bearer " + TokenText, "Bearer " + TokenText }),
                    Is.False);
            }
        }

        [Test]
        public void ReplacementTokenDoesNotAcceptRetiredValue()
        {
            var replacementText = new string('A', 43);
            using (var retired = CreateToken())
            using (var replacement = CreateToken(replacementText))
            {
                Assert.That(retired.MatchesAuthorizationHeaders(new[] { "Bearer " + TokenText }), Is.True);
                Assert.That(replacement.MatchesAuthorizationHeaders(new[] { "Bearer " + TokenText }), Is.False);
                Assert.That(replacement.MatchesAuthorizationHeaders(new[] { "Bearer " + replacementText }), Is.True);
            }
        }

        [Test]
        public void ProcessEnvironmentLoaderRequiresAValidToken()
        {
            var original = Environment.GetEnvironmentVariable(
                BearerToken.EnvironmentVariableName,
                EnvironmentVariableTarget.Process);
            try
            {
                Environment.SetEnvironmentVariable(
                    BearerToken.EnvironmentVariableName,
                    null,
                    EnvironmentVariableTarget.Process);
                Assert.Throws<SecurityException>(
                    (Action)(() => BearerToken.LoadFromProcessEnvironment()));

                Environment.SetEnvironmentVariable(
                    BearerToken.EnvironmentVariableName,
                    TokenText,
                    EnvironmentVariableTarget.Process);
                using (var token = BearerToken.LoadFromProcessEnvironment())
                {
                    Assert.That(
                        token.MatchesAuthorizationHeaders(new[] { "Bearer " + TokenText }),
                        Is.True);
                }
            }
            finally
            {
                Environment.SetEnvironmentVariable(
                    BearerToken.EnvironmentVariableName,
                    original,
                    EnvironmentVariableTarget.Process);
            }
        }

        [Test]
        public void DisposedTokenCannotAuthenticate()
        {
            var token = CreateToken();
            token.Dispose();

            Assert.Throws<ObjectDisposedException>(
                (Action)(() => token.MatchesAuthorizationHeaders(new[] { "Bearer " + TokenText })));
            Assert.Throws<ObjectDisposedException>((Action)(() => token.CreateCursorCodec()));
        }

        [Test]
        public void CursorCodecHasAnIndependentDisposableDerivedKey()
        {
            var token = CreateToken();
            var codec = token.CreateCursorCodec();
            token.Dispose();
            try
            {
                var queryHash = new string('a', HmacCursorCodec.Sha256HexLength);
                var cursor = codec.Encode(new MailboxCursorPayload(
                    queryHash,
                    "Mailbox",
                    "STORE-ID"));

                Assert.That(
                    codec.TryDecode(
                        cursor,
                        HmacCursorKind.ListMailboxes,
                        queryHash,
                        out var decoded),
                    Is.True);
                Assert.That(decoded, Is.TypeOf<MailboxCursorPayload>());
            }
            finally
            {
                codec.Dispose();
            }
        }

        [Test]
        public void DisposingACursorCodecDoesNotRetireItsSourceToken()
        {
            using (var token = CreateToken())
            {
                var codec = token.CreateCursorCodec();
                codec.Dispose();

                Assert.That(
                    token.MatchesAuthorizationHeaders(new[] { "Bearer " + TokenText }),
                    Is.True);
                Assert.Throws<ObjectDisposedException>((Action)(() =>
                    codec.Encode(new MailboxCursorPayload(
                        new string('a', HmacCursorCodec.Sha256HexLength),
                        "Mailbox",
                        "STORE-ID"))));
            }
        }

        private static BearerToken CreateToken(string tokenText = TokenText)
        {
            Assert.That(BearerToken.TryCreate(tokenText, out var token), Is.True);
            return token;
        }
    }
}
