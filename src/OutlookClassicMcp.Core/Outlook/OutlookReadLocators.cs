using System;

namespace OutlookClassicMcp.Core.Outlook
{
    public sealed class MailboxRef
    {
        public const int MaximumStoreIdLength = 4096;

        public MailboxRef(string storeId)
        {
            StoreId = OutlookContractValidation.RequireOpaqueIdentifier(
                storeId,
                MaximumStoreIdLength,
                nameof(storeId));
        }

        public string StoreId { get; }
    }

    public sealed class FolderRef
    {
        public const int MaximumEntryIdLength = 4096;

        public FolderRef(string storeId, string entryId)
        {
            StoreId = OutlookContractValidation.RequireOpaqueIdentifier(
                storeId,
                MailboxRef.MaximumStoreIdLength,
                nameof(storeId));
            EntryId = OutlookContractValidation.RequireOpaqueIdentifier(
                entryId,
                MaximumEntryIdLength,
                nameof(entryId));
        }

        public string StoreId { get; }

        public string EntryId { get; }
    }

    public sealed class ItemRef
    {
        public const int MaximumEntryIdLength = 4096;
        public const int MaximumItemClassLength = 256;

        public ItemRef(string storeId, string entryId, string itemClass)
        {
            StoreId = OutlookContractValidation.RequireOpaqueIdentifier(
                storeId,
                MailboxRef.MaximumStoreIdLength,
                nameof(storeId));
            EntryId = OutlookContractValidation.RequireOpaqueIdentifier(
                entryId,
                MaximumEntryIdLength,
                nameof(entryId));
            ItemClass = OutlookContractValidation.RequireBoundedText(
                itemClass,
                MaximumItemClassLength,
                nameof(itemClass));
        }

        public string StoreId { get; }

        public string EntryId { get; }

        public string ItemClass { get; }
    }

    public sealed class AttachmentRef
    {
        public const int MaximumNameLength = 255;
        public const int FingerprintLength = 64;

        public AttachmentRef(
            ItemRef item,
            int attachmentIndex,
            string name,
            long size,
            bool sizeIsKnown,
            string metadataFingerprint)
        {
            Item = item ?? throw new ArgumentNullException(nameof(item));
            if (attachmentIndex < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(attachmentIndex));
            }

            AttachmentIndex = attachmentIndex;
            Name = OutlookContractValidation.RequireBoundedDisplayText(
                name,
                MaximumNameLength,
                nameof(name));
            if (size < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(size));
            }

            Size = size;
            if (!sizeIsKnown && size != 0)
            {
                throw new ArgumentException(
                    "An unknown attachment size must retain Outlook's zero sentinel.",
                    nameof(sizeIsKnown));
            }

            SizeIsKnown = sizeIsKnown;
            MetadataFingerprint = OutlookContractValidation.RequireSha256Fingerprint(
                metadataFingerprint,
                nameof(metadataFingerprint));
        }

        public ItemRef Item { get; }

        public int AttachmentIndex { get; }

        public string Name { get; }

        public long Size { get; }

        public bool SizeIsKnown { get; }

        public string MetadataFingerprint { get; }
    }

    public sealed class OutlookMailboxKeysetAnchor
    {
        public OutlookMailboxKeysetAnchor(string displayName, MailboxRef mailbox)
        {
            DisplayName = OutlookContractValidation.RequireBoundedDisplayText(
                displayName,
                OutlookMailboxSummary.MaximumDisplayNameLength,
                nameof(displayName));
            Mailbox = mailbox ?? throw new ArgumentNullException(nameof(mailbox));
        }

        public string DisplayName { get; }

        public MailboxRef Mailbox { get; }
    }

    public sealed class OutlookFolderKeysetAnchor
    {
        public OutlookFolderKeysetAnchor(string displayName, FolderRef folder)
        {
            DisplayName = OutlookContractValidation.RequireBoundedDisplayText(
                displayName,
                OutlookFolderSummary.MaximumDisplayNameLength,
                nameof(displayName));
            Folder = folder ?? throw new ArgumentNullException(nameof(folder));
        }

        public string DisplayName { get; }

        public FolderRef Folder { get; }
    }

    public sealed class OutlookMessageKeysetAnchor
    {
        public OutlookMessageKeysetAnchor(DateTime effectiveTimestampUtc, ItemRef item)
        {
            EffectiveTimestampUtc = OutlookContractValidation.RequireUtc(
                effectiveTimestampUtc,
                nameof(effectiveTimestampUtc));
            Item = item ?? throw new ArgumentNullException(nameof(item));
        }

        public DateTime EffectiveTimestampUtc { get; }

        public ItemRef Item { get; }
    }

    public sealed class OutlookAttachmentKeysetAnchor
    {
        public OutlookAttachmentKeysetAnchor(
            int attachmentIndex,
            string metadataFingerprint)
        {
            if (attachmentIndex < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(attachmentIndex));
            }

            AttachmentIndex = attachmentIndex;
            MetadataFingerprint = OutlookContractValidation.RequireSha256Fingerprint(
                metadataFingerprint,
                nameof(metadataFingerprint));
        }

        public int AttachmentIndex { get; }

        public string MetadataFingerprint { get; }
    }
}
