using System;
using System.Collections.Generic;

namespace OutlookClassicMcp.Core.Policy
{
    public static class ToolExposurePolicy
    {
        private static readonly IReadOnlyList<string> NoTools = Array.Empty<string>();
        private static readonly IReadOnlyList<string> StatusOnly =
            Array.AsReadOnly(new[] { ToolNames.OutlookStatus });
        private static readonly IReadOnlyList<string> StatusAndProbe =
            Array.AsReadOnly(new[]
            {
                ToolNames.OutlookStatus,
                ToolNames.OutlookProbe,
            });

        public static IReadOnlyList<string> GetEnabledTools(ImplementationPhase phase)
        {
            if (!Enum.IsDefined(typeof(ImplementationPhase), phase))
            {
                throw new ArgumentOutOfRangeException(nameof(phase));
            }

            if (phase < ImplementationPhase.AuthenticatedTransport)
            {
                return NoTools;
            }

            if (phase < ImplementationPhase.OutlookProbe)
            {
                return StatusOnly;
            }

            return StatusAndProbe;
        }
    }

    public static class ToolNames
    {
        public const string OutlookStatus = "outlook_status";
        public const string OutlookProbe = "outlook_probe";
    }
}
