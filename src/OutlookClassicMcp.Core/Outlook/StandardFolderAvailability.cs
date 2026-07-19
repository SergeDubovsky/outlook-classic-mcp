using System;

namespace OutlookClassicMcp.Core.Outlook
{
    public enum OutlookFolderAvailability
    {
        Unknown = 0,
        Available = 1,
        Missing = 2,
    }

    public sealed class StandardFolderAvailability
    {
        public StandardFolderAvailability(
            OutlookFolderAvailability inbox,
            OutlookFolderAvailability drafts,
            OutlookFolderAvailability sent,
            OutlookFolderAvailability deleted,
            OutlookFolderAvailability archive)
        {
            OutlookContractValidation.RequireDefinedEnum(inbox, nameof(inbox));
            OutlookContractValidation.RequireDefinedEnum(drafts, nameof(drafts));
            OutlookContractValidation.RequireDefinedEnum(sent, nameof(sent));
            OutlookContractValidation.RequireDefinedEnum(deleted, nameof(deleted));
            OutlookContractValidation.RequireDefinedEnum(archive, nameof(archive));

            Inbox = inbox;
            Drafts = drafts;
            Sent = sent;
            Deleted = deleted;
            Archive = archive;
        }

        public OutlookFolderAvailability Inbox { get; }

        public OutlookFolderAvailability Drafts { get; }

        public OutlookFolderAvailability Sent { get; }

        public OutlookFolderAvailability Deleted { get; }

        public OutlookFolderAvailability Archive { get; }
    }
}
