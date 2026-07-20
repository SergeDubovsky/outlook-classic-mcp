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
        private static readonly IReadOnlyList<string> BoundedReadTools =
            Array.AsReadOnly(new[]
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

            if (phase < ImplementationPhase.BoundedReads)
            {
                return StatusAndProbe;
            }

            return BoundedReadTools;
        }
    }

    public static class ToolNames
    {
        public const string OutlookStatus = "outlook_status";
        public const string OutlookProbe = "outlook_probe";
        public const string OutlookListMailboxes = "outlook_list_mailboxes";
        public const string OutlookListFolders = "outlook_list_folders";
        public const string OutlookListMessages = "outlook_list_messages";
        public const string OutlookSearchMessages = "outlook_search_messages";
        public const string OutlookGetMessage = "outlook_get_message";
        public const string OutlookGetConversation = "outlook_get_conversation";
        public const string OutlookListAttachments = "outlook_list_attachments";
    }
}
