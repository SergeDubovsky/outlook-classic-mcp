using NUnit.Framework;
using OutlookClassicMcp.AddIn.Runtime;

namespace OutlookClassicMcp.Core.Tests
{
    [TestFixture]
    public sealed class HostLifecycleTests
    {
        [Test]
        public void StartupMovesCreatedHostOnline()
        {
            var lifecycle = new HostLifecycle();

            Assert.That(lifecycle.State, Is.EqualTo(HostLifecycleState.Created));
            Assert.That(lifecycle.TryBeginStartup(), Is.True);
            Assert.That(lifecycle.State, Is.EqualTo(HostLifecycleState.Starting));

            Assert.That(lifecycle.TryMarkOnline(), Is.True);

            Assert.That(lifecycle.State, Is.EqualTo(HostLifecycleState.Online));
            Assert.That(lifecycle.TryBeginStartup(), Is.False);
        }

        [Test]
        public void DegradedHostCanRetryStartup()
        {
            var lifecycle = new HostLifecycle();
            lifecycle.TryBeginStartup();
            Assert.That(lifecycle.TryMarkDegraded(), Is.True);

            Assert.That(lifecycle.TryBeginStartup(), Is.True);
            Assert.That(lifecycle.TryMarkOnline(), Is.True);

            Assert.That(lifecycle.State, Is.EqualTo(HostLifecycleState.Online));
        }

        [Test]
        public void OnlineHostCanFailClosedWhenItsListenerStops()
        {
            var lifecycle = CreateOnlineLifecycle();

            Assert.That(lifecycle.TryMarkDegraded(), Is.True);

            Assert.That(lifecycle.State, Is.EqualTo(HostLifecycleState.Degraded));
            Assert.That(lifecycle.TryMarkDegraded(), Is.False);
        }

        [Test]
        public void PausedHostCanResumeThroughStartup()
        {
            var lifecycle = CreateOnlineLifecycle();

            Assert.That(lifecycle.TryBeginPause(), Is.True);
            Assert.That(lifecycle.TryMarkPaused(), Is.True);
            Assert.That(lifecycle.State, Is.EqualTo(HostLifecycleState.Paused));

            Assert.That(lifecycle.TryBeginStartup(), Is.True);
            Assert.That(lifecycle.TryMarkOnline(), Is.True);

            Assert.That(lifecycle.State, Is.EqualTo(HostLifecycleState.Online));
        }

        [Test]
        public void StopIsIdempotent()
        {
            var lifecycle = CreateOnlineLifecycle();

            Assert.That(lifecycle.TryBeginStop(), Is.True);
            Assert.That(lifecycle.TryBeginStop(), Is.False);
            lifecycle.MarkStopped();
            lifecycle.MarkStopped();

            Assert.That(lifecycle.State, Is.EqualTo(HostLifecycleState.Stopped));
            Assert.That(lifecycle.TryBeginStop(), Is.False);
        }

        [Test]
        public void InvalidCompletionTransitionFailsClosed()
        {
            var lifecycle = new HostLifecycle();

            Assert.That(lifecycle.TryMarkOnline(), Is.False);
            Assert.That(lifecycle.State, Is.EqualTo(HostLifecycleState.Created));
        }

        [Test]
        public void ShutdownDuringStartupCannotResurrectTheHost()
        {
            var lifecycle = new HostLifecycle();
            Assert.That(lifecycle.TryBeginStartup(), Is.True);

            Assert.That(lifecycle.TryBeginStop(), Is.True);
            Assert.That(lifecycle.TryMarkOnline(), Is.False);
            Assert.That(lifecycle.TryMarkDegraded(), Is.False);
            lifecycle.MarkStopped();

            Assert.That(lifecycle.State, Is.EqualTo(HostLifecycleState.Stopped));
        }

        [Test]
        public void ShutdownCanBeginFromEveryNonterminalState()
        {
            foreach (var lifecycle in CreateNonterminalLifecycles())
            {
                Assert.That(lifecycle.TryBeginStop(), Is.True);
                lifecycle.MarkStopped();
                Assert.That(lifecycle.State, Is.EqualTo(HostLifecycleState.Stopped));
            }
        }

        private static HostLifecycle CreateOnlineLifecycle()
        {
            var lifecycle = new HostLifecycle();
            lifecycle.TryBeginStartup();
            Assert.That(lifecycle.TryMarkOnline(), Is.True);
            return lifecycle;
        }

        private static HostLifecycle[] CreateNonterminalLifecycles()
        {
            var created = new HostLifecycle();

            var starting = new HostLifecycle();
            starting.TryBeginStartup();

            var online = CreateOnlineLifecycle();

            var degraded = new HostLifecycle();
            degraded.TryBeginStartup();
            degraded.TryMarkDegraded();

            var pausing = CreateOnlineLifecycle();
            pausing.TryBeginPause();

            var paused = CreateOnlineLifecycle();
            paused.TryBeginPause();
            paused.TryMarkPaused();

            return new[] { created, starting, online, degraded, pausing, paused };
        }
    }
}
