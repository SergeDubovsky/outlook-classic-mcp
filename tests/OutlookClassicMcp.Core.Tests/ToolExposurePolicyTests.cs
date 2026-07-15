using System;
using NUnit.Framework;
using OutlookClassicMcp.Core.Policy;

namespace OutlookClassicMcp.Core.Tests
{
    [TestFixture]
    public sealed class ToolExposurePolicyTests
    {
        [Test]
        public void PhaseZeroExposesNoTools()
        {
            var tools = ToolExposurePolicy.GetEnabledTools(ImplementationPhase.RepositoryAndToolchain);

            Assert.That(tools, Is.Empty);
        }

        [Test]
        public void StatusAppearsOnlyAtAuthenticatedTransportPhase()
        {
            var lifecycleTools = ToolExposurePolicy.GetEnabledTools(ImplementationPhase.DependencyAndLifecycle);
            var transportTools = ToolExposurePolicy.GetEnabledTools(ImplementationPhase.AuthenticatedTransport);

            Assert.That(lifecycleTools, Is.Empty);
            Assert.That(transportTools, Is.EqualTo(new[] { ToolNames.OutlookStatus }));
        }

        [Test]
        public void UnknownPhaseFailsClosed()
        {
            Assert.Throws<ArgumentOutOfRangeException>(
                (Action)(() => ToolExposurePolicy.GetEnabledTools((ImplementationPhase)99)));
        }
    }
}
