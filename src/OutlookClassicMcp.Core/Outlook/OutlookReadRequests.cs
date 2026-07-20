using System;
using System.Collections.Generic;

namespace OutlookClassicMcp.Core.Outlook
{
    public static class OutlookReadLimits
    {
        public const int MinimumPageSize = 1;
        public const int MaximumPageSize = 50;
        public const int MaximumSearchScopeCount = 64;
        public const int MaximumBodyCharacters = 50000;
    }

    public sealed class OutlookListMailboxesRequest
    {
        public OutlookListMailboxesRequest(int pageSize, OutlookMailboxKeysetAnchor? anchor)
        {
            PageSize = ValidatePageSize(pageSize);
            Anchor = anchor;
        }

        public int PageSize { get; }

        public OutlookMailboxKeysetAnchor? Anchor { get; }

        internal static int ValidatePageSize(int pageSize)
        {
            if (pageSize < OutlookReadLimits.MinimumPageSize ||
                pageSize > OutlookReadLimits.MaximumPageSize)
            {
                throw new ArgumentOutOfRangeException(nameof(pageSize));
            }

            return pageSize;
        }
    }

    public sealed class OutlookListFoldersRequest
    {
        public OutlookListFoldersRequest(
            MailboxRef mailbox,
            FolderRef? parentFolder,
            int pageSize,
            OutlookFolderKeysetAnchor? anchor)
        {
            Mailbox = mailbox ?? throw new ArgumentNullException(nameof(mailbox));
            if (parentFolder != null &&
                !string.Equals(mailbox.StoreId, parentFolder.StoreId, StringComparison.Ordinal))
            {
                throw new ArgumentException(
                    "The parent folder must belong to the selected mailbox.",
                    nameof(parentFolder));
            }

            if (anchor != null &&
                !string.Equals(mailbox.StoreId, anchor.Folder.StoreId, StringComparison.Ordinal))
            {
                throw new ArgumentException(
                    "The anchor must belong to the selected mailbox.",
                    nameof(anchor));
            }

            ParentFolder = parentFolder;
            PageSize = OutlookListMailboxesRequest.ValidatePageSize(pageSize);
            Anchor = anchor;
        }

        public MailboxRef Mailbox { get; }

        public FolderRef? ParentFolder { get; }

        public int PageSize { get; }

        public OutlookFolderKeysetAnchor? Anchor { get; }
    }

    public sealed class OutlookListMessagesRequest
    {
        public OutlookListMessagesRequest(
            FolderRef folder,
            int pageSize,
            OutlookMessageKeysetAnchor? anchor)
        {
            Folder = folder ?? throw new ArgumentNullException(nameof(folder));
            if (anchor != null &&
                !string.Equals(folder.StoreId, anchor.Item.StoreId, StringComparison.Ordinal))
            {
                throw new ArgumentException(
                    "The anchor must belong to the selected folder's store.",
                    nameof(anchor));
            }

            PageSize = OutlookListMailboxesRequest.ValidatePageSize(pageSize);
            Anchor = anchor;
        }

        public FolderRef Folder { get; }

        public int PageSize { get; }

        public OutlookMessageKeysetAnchor? Anchor { get; }
    }

    public sealed class OutlookSearchScope
    {
        public OutlookSearchScope(MailboxRef mailbox, FolderRef? folder)
        {
            Mailbox = mailbox ?? throw new ArgumentNullException(nameof(mailbox));
            if (folder != null &&
                !string.Equals(mailbox.StoreId, folder.StoreId, StringComparison.Ordinal))
            {
                throw new ArgumentException(
                    "The folder must belong to the selected mailbox.",
                    nameof(folder));
            }

            Folder = folder;
        }

        public MailboxRef Mailbox { get; }

        public FolderRef? Folder { get; }

        internal string Key => Mailbox.StoreId + "\n" + (Folder?.EntryId ?? string.Empty);
    }

    public sealed class OutlookMessageSearchFilter
    {
        public const int MaximumAddressLength = 512;
        public const int MaximumSubjectLength = 1024;
        public const int MaximumTextLength = 4096;
        public const int MaximumCategoryLength = 256;

        public OutlookMessageSearchFilter(
            string? sender,
            string? recipient,
            string? subject,
            string? text,
            DateTime? receivedFromUtc,
            DateTime? receivedToUtc,
            bool? isUnread,
            string? category,
            bool? hasAttachments)
        {
            Sender = OutlookContractValidation.OptionalBoundedText(
                sender,
                MaximumAddressLength,
                nameof(sender));
            Recipient = OutlookContractValidation.OptionalBoundedText(
                recipient,
                MaximumAddressLength,
                nameof(recipient));
            Subject = OutlookContractValidation.OptionalBoundedText(
                subject,
                MaximumSubjectLength,
                nameof(subject));
            Text = OutlookContractValidation.OptionalBoundedText(
                text,
                MaximumTextLength,
                nameof(text));
            ReceivedFromUtc = OutlookContractValidation.OptionalUtc(
                receivedFromUtc,
                nameof(receivedFromUtc));
            ReceivedToUtc = OutlookContractValidation.OptionalUtc(
                receivedToUtc,
                nameof(receivedToUtc));
            if (ReceivedFromUtc > ReceivedToUtc)
            {
                throw new ArgumentException(
                    "The received range start must not follow its end.",
                    nameof(receivedFromUtc));
            }

            IsUnread = isUnread;
            Category = OutlookContractValidation.OptionalBoundedText(
                category,
                MaximumCategoryLength,
                nameof(category));
            HasAttachments = hasAttachments;
        }

        public string? Sender { get; }

        public string? Recipient { get; }

        public string? Subject { get; }

        public string? Text { get; }

        public DateTime? ReceivedFromUtc { get; }

        public DateTime? ReceivedToUtc { get; }

        public bool? IsUnread { get; }

        public string? Category { get; }

        public bool? HasAttachments { get; }
    }

    public sealed class OutlookSearchMessagesRequest
    {
        public OutlookSearchMessagesRequest(
            IEnumerable<OutlookSearchScope> scopes,
            OutlookMessageSearchFilter filter,
            int pageSize,
            OutlookMessageKeysetAnchor? anchor)
        {
            var scopeCopy = OutlookContractValidation.BoundedCopy(
                scopes,
                OutlookReadLimits.MaximumSearchScopeCount,
                nameof(scopes),
                rejectNullElements: true);
            if (scopeCopy.Count == 0)
            {
                throw new ArgumentException(
                    "At least one explicit search scope is required.",
                    nameof(scopes));
            }

            var keys = new HashSet<string>(StringComparer.Ordinal);
            foreach (var scope in scopeCopy)
            {
                if (!keys.Add(scope.Key))
                {
                    throw new ArgumentException(
                        "Search scopes must be distinct.",
                        nameof(scopes));
                }
            }

            if (anchor != null && !keys.ContainsStore(anchor.Item.StoreId))
            {
                throw new ArgumentException(
                    "The anchor must belong to a selected search store.",
                    nameof(anchor));
            }

            Scopes = scopeCopy;
            Filter = filter ?? throw new ArgumentNullException(nameof(filter));
            PageSize = OutlookListMailboxesRequest.ValidatePageSize(pageSize);
            Anchor = anchor;
        }

        public IReadOnlyList<OutlookSearchScope> Scopes { get; }

        public OutlookMessageSearchFilter Filter { get; }

        public int PageSize { get; }

        public OutlookMessageKeysetAnchor? Anchor { get; }
    }

    public enum OutlookBodyFormat
    {
        PlainText = 0,
        Html = 1,
    }

    public sealed class OutlookGetMessageRequest
    {
        public OutlookGetMessageRequest(
            ItemRef item,
            OutlookBodyFormat bodyFormat,
            int maximumBodyCharacters)
        {
            Item = item ?? throw new ArgumentNullException(nameof(item));
            OutlookContractValidation.RequireDefinedEnum(bodyFormat, nameof(bodyFormat));
            BodyFormat = bodyFormat;
            if (maximumBodyCharacters < 1 ||
                maximumBodyCharacters > OutlookReadLimits.MaximumBodyCharacters)
            {
                throw new ArgumentOutOfRangeException(nameof(maximumBodyCharacters));
            }

            MaximumBodyCharacters = maximumBodyCharacters;
        }

        public ItemRef Item { get; }

        public OutlookBodyFormat BodyFormat { get; }

        public int MaximumBodyCharacters { get; }
    }

    public sealed class OutlookGetConversationRequest
    {
        public OutlookGetConversationRequest(
            ItemRef item,
            int pageSize,
            OutlookMessageKeysetAnchor? anchor)
        {
            Item = item ?? throw new ArgumentNullException(nameof(item));
            PageSize = OutlookListMailboxesRequest.ValidatePageSize(pageSize);
            Anchor = anchor;
        }

        public ItemRef Item { get; }

        public int PageSize { get; }

        public OutlookMessageKeysetAnchor? Anchor { get; }
    }

    public sealed class OutlookListAttachmentsRequest
    {
        public OutlookListAttachmentsRequest(
            ItemRef item,
            int pageSize,
            OutlookAttachmentKeysetAnchor? anchor)
        {
            Item = item ?? throw new ArgumentNullException(nameof(item));
            PageSize = OutlookListMailboxesRequest.ValidatePageSize(pageSize);
            Anchor = anchor;
        }

        public ItemRef Item { get; }

        public int PageSize { get; }

        public OutlookAttachmentKeysetAnchor? Anchor { get; }
    }

    internal static class OutlookSearchScopeKeySetExtensions
    {
        public static bool ContainsStore(this ISet<string> keys, string storeId)
        {
            foreach (var key in keys)
            {
                if (key.StartsWith(storeId + "\n", StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }
    }
}
