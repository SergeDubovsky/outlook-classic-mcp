using System;
using NUnit.Framework;

namespace OutlookClassicMcp.Transport.Tests
{
    [TestFixture]
    [NonParallelizable]
    public sealed class LoopbackHttpServerTests
    {
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
            using (var first = new LoopbackHttpServer())
            {
                first.Start();
                Assert.That(first.IsListening, Is.True);
                first.Stop();
                Assert.That(first.IsListening, Is.False);
                Assert.DoesNotThrow((Action)first.Stop);

                using (var second = new LoopbackHttpServer())
                {
                    Assert.DoesNotThrow((Action)second.Start);
                    Assert.That(second.IsListening, Is.True);
                }
            }
        }

        [Test]
        public void DisposedListenerCannotRestart()
        {
            var server = new LoopbackHttpServer();
            server.Dispose();

            Assert.Throws<ObjectDisposedException>((Action)server.Start);
        }
    }
}
