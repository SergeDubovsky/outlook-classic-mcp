using System;
using System.Linq;
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

        [Test]
        public void BoundedReadsExposePriorToolsPlusExactlySevenReadTools()
        {
            var tools = ToolExposurePolicy.GetEnabledTools(ImplementationPhase.BoundedReads);

            Assert.That(
                tools,
                Is.EqualTo(new[]
                {
                    ToolNames.OutlookStatus,
                    ToolNames.OutlookProbe,
                    ToolNames.OutlookListMailboxes,
                    ToolNames.OutlookListFolders,
                    ToolNames.OutlookListMessages,
                    ToolNames.OutlookSearchMessages,
                    ToolNames.OutlookGetMessage,
                    ToolNames.OutlookGetConversation,
                    ToolNames.OutlookListAttachments,
                }));
            Assert.That(tools.Distinct().ToArray(), Has.Length.EqualTo(9));
        }
    }
}
