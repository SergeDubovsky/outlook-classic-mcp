using System;
using System.Collections.Generic;

namespace OutlookClassicMcp.Core.Outlook
{
    public sealed class OutlookStandardFolderReferences
    {
        public OutlookStandardFolderReferences(
            FolderRef? inbox,
            FolderRef? drafts,
            FolderRef? sent,
            FolderRef? deleted,
            FolderRef? archive)
        {
            Inbox = inbox;
            Drafts = drafts;
            Sent = sent;
            Deleted = deleted;
            Archive = archive;
        }

        public FolderRef? Inbox { get; }

        public FolderRef? Drafts { get; }

        public FolderRef? Sent { get; }

        public FolderRef? Deleted { get; }

        public FolderRef? Archive { get; }

        internal IEnumerable<FolderRef?> Enumerate()
        {
            yield return Inbox;
            yield return Drafts;
            yield return Sent;
            yield return Deleted;
            yield return Archive;
        }
    }

    public sealed class OutlookMailboxSummary
    {
        public const int MaximumDisplayNameLength = 256;

        public OutlookMailboxSummary(
            MailboxRef mailbox,
            string displayName,
            OutlookStoreType storeType,
            OutlookStoreCapabilities capabilities,
            OutlookStandardFolderReferences standardFolders)
        {
            Mailbox = mailbox ?? throw new ArgumentNullException(nameof(mailbox));
            DisplayName = OutlookContractValidation.RequireBoundedDisplayText(
                displayName,
                MaximumDisplayNameLength,
                nameof(displayName));
            OutlookContractValidation.RequireDefinedEnum(storeType, nameof(storeType));
            StoreType = storeType;
            Capabilities = capabilities ?? throw new ArgumentNullException(nameof(capabilities));
            StandardFolders = standardFolders ?? throw new ArgumentNullException(nameof(standardFolders));
            foreach (var folder in standardFolders.Enumerate())
            {
                if (folder != null &&
                    !string.Equals(mailbox.StoreId, folder.StoreId, StringComparison.Ordinal))
                {
                    throw new ArgumentException(
                        "Every standard folder must belong to the mailbox.",
                        nameof(standardFolders));
                }
            }
        }

        public MailboxRef Mailbox { get; }

        public string DisplayName { get; }

        public OutlookStoreType StoreType { get; }

        public OutlookStoreCapabilities Capabilities { get; }

        public OutlookStandardFolderReferences StandardFolders { get; }
    }

    public sealed class OutlookFolderSummary
    {
        public const int MaximumDisplayNameLength = 256;

        public OutlookFolderSummary(
            FolderRef folder,
            FolderRef? parentFolder,
            string displayName,
            bool hasChildren)
        {
            Folder = folder ?? throw new ArgumentNullException(nameof(folder));
            if (parentFolder != null &&
                !string.Equals(folder.StoreId, parentFolder.StoreId, StringComparison.Ordinal))
            {
                throw new ArgumentException(
                    "The parent folder must belong to the same store.",
                    nameof(parentFolder));
            }

            ParentFolder = parentFolder;
            DisplayName = OutlookContractValidation.RequireBoundedDisplayText(
                displayName,
                MaximumDisplayNameLength,
                nameof(displayName));
            HasChildren = hasChildren;
        }

        public FolderRef Folder { get; }

        public FolderRef? ParentFolder { get; }

        public string DisplayName { get; }

        public bool HasChildren { get; }
    }

    public sealed class OutlookMessageSummary
    {
        public const int MaximumSubjectLength = 1024;
        public const int MaximumSenderLength = 512;
        public const int MaximumConversationIdLength = 4096;
        public OutlookMessageSummary(
            ItemRef item,
            FolderRef folder,
            string subject,
            string? senderDisplayName,
            string? senderAddress,
            DateTime effectiveTimestampUtc,
            DateTime? receivedUtc,
            DateTime? sentUtc,
            bool isRead,
            int attachmentCount,
            string? conversationId)
        {
            Item = item ?? throw new ArgumentNullException(nameof(item));
            Folder = folder ?? throw new ArgumentNullException(nameof(folder));
            if (!string.Equals(item.StoreId, folder.StoreId, StringComparison.Ordinal))
            {
                throw new ArgumentException(
                    "The item and folder must belong to the same store.",
                    nameof(folder));
            }

            Subject = OutlookContractValidation.RequireBoundedDisplayText(
                subject,
                MaximumSubjectLength,
                nameof(subject));
            SenderDisplayName = senderDisplayName == null
                ? null
                : OutlookContractValidation.RequireBoundedDisplayText(
                    senderDisplayName,
                    MaximumSenderLength,
                    nameof(senderDisplayName));
            SenderAddress = OutlookContractValidation.OptionalBoundedText(
                senderAddress,
                MaximumSenderLength,
                nameof(senderAddress));
            EffectiveTimestampUtc = OutlookContractValidation.RequireUtc(
                effectiveTimestampUtc,
                nameof(effectiveTimestampUtc));
            ReceivedUtc = OutlookContractValidation.OptionalUtc(receivedUtc, nameof(receivedUtc));
            SentUtc = OutlookContractValidation.OptionalUtc(sentUtc, nameof(sentUtc));
            IsRead = isRead;
            if (attachmentCount < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(attachmentCount));
            }

            AttachmentCount = attachmentCount;
            ConversationId = conversationId == null
                ? null
                : OutlookContractValidation.RequireOpaqueIdentifier(
                    conversationId,
                    MaximumConversationIdLength,
                    nameof(conversationId));
        }

        public ItemRef Item { get; }

        public FolderRef Folder { get; }

        public string Subject { get; }

        public string? SenderDisplayName { get; }

        public string? SenderAddress { get; }

        public DateTime EffectiveTimestampUtc { get; }

        public DateTime? ReceivedUtc { get; }

        public DateTime? SentUtc { get; }

        public bool IsRead { get; }

        public int AttachmentCount { get; }

        public bool HasAttachments => AttachmentCount > 0;

        public string? ConversationId { get; }
    }

    public sealed class OutlookMessageAddress
    {
        public const int MaximumValueLength = 512;

        public OutlookMessageAddress(string? displayName, string? address)
        {
            DisplayName = displayName == null
                ? null
                : OutlookContractValidation.RequireBoundedDisplayText(
                    displayName,
                    MaximumValueLength,
                    nameof(displayName));
            Address = OutlookContractValidation.OptionalBoundedText(
                address,
                MaximumValueLength,
                nameof(address));
            if (DisplayName == null && Address == null)
            {
                throw new ArgumentException("A display name or address is required.");
            }
        }

        public string? DisplayName { get; }

        public string? Address { get; }
    }

    public sealed class OutlookMessageBody
    {
        public OutlookMessageBody(
            OutlookBodyFormat format,
            string content,
            int? originalCharacterCount,
            bool isTruncated,
            bool isProtected)
        {
            OutlookContractValidation.RequireDefinedEnum(format, nameof(format));
            Format = format;
            Content = OutlookContractValidation.RequireBoundedContent(
                content,
                OutlookReadLimits.MaximumBodyCharacters,
                nameof(content));

            if (isProtected)
            {
                if (content.Length != 0 || originalCharacterCount.HasValue || isTruncated)
                {
                    throw new ArgumentException(
                        "Protected content must not include body data or truncation metadata.",
                        nameof(isProtected));
                }
            }
            else
            {
                if (!originalCharacterCount.HasValue ||
                    originalCharacterCount.Value < content.Length ||
                    isTruncated != (originalCharacterCount.Value > content.Length))
                {
                    throw new ArgumentException(
                        "Body length and truncation metadata are inconsistent.",
                        nameof(originalCharacterCount));
                }
            }

            OriginalCharacterCount = originalCharacterCount;
            IsTruncated = isTruncated;
            IsProtected = isProtected;
        }

        public OutlookBodyFormat Format { get; }

        public string Content { get; }

        public int? OriginalCharacterCount { get; }

        public bool IsTruncated { get; }

        public bool IsProtected { get; }
    }

    public sealed class OutlookMessageDetail
    {
        public const int MaximumRecipientCount = 128;

        public OutlookMessageDetail(
            OutlookMessageSummary summary,
            IEnumerable<OutlookMessageAddress> toRecipients,
            IEnumerable<OutlookMessageAddress> ccRecipients,
            IEnumerable<OutlookMessageAddress> bccRecipients,
            int totalToRecipientCount,
            int totalCcRecipientCount,
            int totalBccRecipientCount,
            OutlookMessageBody body)
        {
            Summary = summary ?? throw new ArgumentNullException(nameof(summary));
            ToRecipients = OutlookContractValidation.BoundedCopy(
                toRecipients,
                MaximumRecipientCount,
                nameof(toRecipients),
                rejectNullElements: true);
            CcRecipients = OutlookContractValidation.BoundedCopy(
                ccRecipients,
                MaximumRecipientCount,
                nameof(ccRecipients),
                rejectNullElements: true);
            BccRecipients = OutlookContractValidation.BoundedCopy(
                bccRecipients,
                MaximumRecipientCount,
                nameof(bccRecipients),
                rejectNullElements: true);
            TotalToRecipientCount = ValidateTotalRecipientCount(
                totalToRecipientCount,
                ToRecipients.Count,
                nameof(totalToRecipientCount));
            TotalCcRecipientCount = ValidateTotalRecipientCount(
                totalCcRecipientCount,
                CcRecipients.Count,
                nameof(totalCcRecipientCount));
            TotalBccRecipientCount = ValidateTotalRecipientCount(
                totalBccRecipientCount,
                BccRecipients.Count,
                nameof(totalBccRecipientCount));
            Body = body ?? throw new ArgumentNullException(nameof(body));
        }

        public OutlookMessageSummary Summary { get; }

        public IReadOnlyList<OutlookMessageAddress> ToRecipients { get; }

        public IReadOnlyList<OutlookMessageAddress> CcRecipients { get; }

        public IReadOnlyList<OutlookMessageAddress> BccRecipients { get; }

        public int TotalToRecipientCount { get; }

        public int TotalCcRecipientCount { get; }

        public int TotalBccRecipientCount { get; }

        public bool ToRecipientsTruncated => TotalToRecipientCount > ToRecipients.Count;

        public bool CcRecipientsTruncated => TotalCcRecipientCount > CcRecipients.Count;

        public bool BccRecipientsTruncated => TotalBccRecipientCount > BccRecipients.Count;

        public OutlookMessageBody Body { get; }

        private static int ValidateTotalRecipientCount(
            int totalCount,
            int returnedCount,
            string parameterName)
        {
            if (totalCount < returnedCount)
            {
                throw new ArgumentOutOfRangeException(
                    parameterName,
                    "The total recipient count cannot be smaller than the returned count.");
            }

            return totalCount;
        }
    }

    public sealed class OutlookAttachmentSummary
    {
        public const int MaximumContentTypeLength = 256;

        public OutlookAttachmentSummary(AttachmentRef attachment, string? contentType)
        {
            Attachment = attachment ?? throw new ArgumentNullException(nameof(attachment));
            ContentType = OutlookContractValidation.OptionalBoundedText(
                contentType,
                MaximumContentTypeLength,
                nameof(contentType));
        }

        public AttachmentRef Attachment { get; }

        public string? ContentType { get; }
    }

    public sealed class OutlookScopeFailure
    {
        public OutlookScopeFailure(OutlookSearchScope scope, OutlookGatewayFailure failure)
        {
            Scope = scope ?? throw new ArgumentNullException(nameof(scope));
            OutlookContractValidation.RequireDefinedEnum(failure, nameof(failure));
            if (!IsScopedFailure(failure))
            {
                throw new ArgumentException(
                    "Only a scope-local failure can be represented as a partial result.",
                    nameof(failure));
            }

            Failure = failure;
        }

        public OutlookSearchScope Scope { get; }

        public OutlookGatewayFailure Failure { get; }

        private static bool IsScopedFailure(OutlookGatewayFailure failure)
        {
            switch (failure)
            {
                case OutlookGatewayFailure.StoreNotFound:
                case OutlookGatewayFailure.FolderNotFound:
                case OutlookGatewayFailure.UnsupportedStore:
                case OutlookGatewayFailure.UnsupportedItemType:
                case OutlookGatewayFailure.AccessDenied:
                case OutlookGatewayFailure.ObjectModelGuard:
                case OutlookGatewayFailure.Timeout:
                case OutlookGatewayFailure.ComBusy:
                    return true;
                default:
                    return false;
            }
        }
    }
}
