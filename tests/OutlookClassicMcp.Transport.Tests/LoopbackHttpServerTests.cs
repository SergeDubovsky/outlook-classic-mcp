using System;
using NUnit.Framework;

namespace OutlookClassicMcp.Transport.Tests
{
    [TestFixture]
    [NonParallelizable]
    public sealed class LoopbackHttpServerTests
    {
        private const string TokenText = "AAECAwQFBgcICQoLDA0ODxAREhMUFRYXGBkaGxwdHh8";
        private static readonly string[] ValidAuthorizationHeaders =
            { "Bearer " + TokenText };

        [Test]
        public void EndpointIsTheExactCanonicalLoopbackPrefix()
        {
            Assert.That(LoopbackEndpoint.Prefix, Is.EqualTo("http://127.0.0.1:8765/mcp/"));
            Assert.That(LoopbackEndpoint.Address.Host, Is.EqualTo("127.0.0.1"));
            Assert.That(LoopbackEndpoint.Address.Port, Is.EqualTo(8765));
            Assert.That(LoopbackEndpoint.Address.AbsolutePath, Is.EqualTo("/mcp/"));
        }

        [Test]
        public void ListenerStartsAndReleasesTheExactPrefix()
        {
            using (var first = CreateServer())
            {
                first.Start();
                Assert.That(first.IsListening, Is.True);
                first.Stop();
                Assert.That(first.IsListening, Is.False);
                Assert.DoesNotThrow((Action)first.Stop);

                using (var second = CreateServer())
                {
                    Assert.DoesNotThrow((Action)second.Start);
                    Assert.That(second.IsListening, Is.True);
                }
            }
        }

        [Test]
        public void DisposedListenerCannotRestart()
        {
            var server = CreateServer();
            server.Dispose();

            Assert.Throws<ObjectDisposedException>((Action)server.Start);
        }

        [Test]
        public void ToolDeadlineLeavesResponseGraceBeforeTheHandlerDeadline()
        {
            Assert.That(RequestLimits.DefaultToolDeadline,
                Is.LessThan(RequestLimits.DefaultHandlerDeadline));
            Assert.That(BearerToken.TryCreate(TokenText, out var token), Is.True);
            using (token)
            {
                Assert.Throws<ArgumentOutOfRangeException>((Action)(() =>
                    _ = new LoopbackHttpServer(
                        token,
                        () => new OutlookStatusSnapshot("online", true, "1.0.0"),
                        new FakeOutlookGateway(),
                        toolDeadline: TimeSpan.FromSeconds(2),
                        handlerDeadline: TimeSpan.FromSeconds(2))));
            }
        }

        [Test]
        public void ConstructorFailureLeavesCallerTokenUsable()
        {
            Assert.That(BearerToken.TryCreate(TokenText, out var token), Is.True);
            using (token)
            {
                var exception = Assert.Throws<ArgumentNullException>((Action)(() =>
                    _ = new LoopbackHttpServer(
                        token,
                        null!,
                        new FakeOutlookGateway())));
                Assert.That(exception!.ParamName, Is.EqualTo("statusProvider"));
                Assert.That(
                    token.MatchesAuthorizationHeaders(ValidAuthorizationHeaders),
                    Is.True);
            }
        }

        private static LoopbackHttpServer CreateServer()
        {
            Assert.That(BearerToken.TryCreate(TokenText, out var token), Is.True);
            return new LoopbackHttpServer(
                token,
                () => new OutlookStatusSnapshot("online", true, "1.0.0"),
                new FakeOutlookGateway());
        }
    }
}
