using System;
using System.Collections.Generic;

namespace OutlookClassicMcp.Core.Outlook
{
    public sealed class OutlookMailboxPage
    {
        public OutlookMailboxPage(
            IEnumerable<OutlookMailboxSummary> items,
            OutlookMailboxKeysetAnchor? nextAnchor)
        {
            var itemCopy = OutlookContractValidation.BoundedCopy(
                items,
                OutlookReadLimits.MaximumPageSize,
                nameof(items),
                rejectNullElements: true);
            ValidateNextAnchor(itemCopy, nextAnchor);
            Items = itemCopy;
            NextAnchor = nextAnchor;
        }

        public IReadOnlyList<OutlookMailboxSummary> Items { get; }

        public OutlookMailboxKeysetAnchor? NextAnchor { get; }

        private static void ValidateNextAnchor(
            IReadOnlyList<OutlookMailboxSummary> items,
            OutlookMailboxKeysetAnchor? nextAnchor)
        {
            if (nextAnchor == null)
            {
                return;
            }

            if (items.Count == 0)
            {
                throw new ArgumentException(
                    "An empty page cannot have a continuation anchor.",
                    nameof(nextAnchor));
            }

            var last = items[items.Count - 1];
            if (!string.Equals(last.DisplayName, nextAnchor.DisplayName, StringComparison.Ordinal) ||
                !string.Equals(last.Mailbox.StoreId, nextAnchor.Mailbox.StoreId, StringComparison.Ordinal))
            {
                throw new ArgumentException(
                    "The continuation anchor must identify the final page item.",
                    nameof(nextAnchor));
            }
        }
    }

    public sealed class OutlookFolderPage
    {
        public OutlookFolderPage(
            IEnumerable<OutlookFolderSummary> items,
            OutlookFolderKeysetAnchor? nextAnchor)
        {
            var itemCopy = OutlookContractValidation.BoundedCopy(
                items,
                OutlookReadLimits.MaximumPageSize,
                nameof(items),
                rejectNullElements: true);
            ValidateNextAnchor(itemCopy, nextAnchor);
            Items = itemCopy;
            NextAnchor = nextAnchor;
        }

        public IReadOnlyList<OutlookFolderSummary> Items { get; }

        public OutlookFolderKeysetAnchor? NextAnchor { get; }

        private static void ValidateNextAnchor(
            IReadOnlyList<OutlookFolderSummary> items,
            OutlookFolderKeysetAnchor? nextAnchor)
        {
            if (nextAnchor == null)
            {
                return;
            }

            if (items.Count == 0)
            {
                throw new ArgumentException(
                    "An empty page cannot have a continuation anchor.",
                    nameof(nextAnchor));
            }

            var last = items[items.Count - 1];
            if (!string.Equals(last.DisplayName, nextAnchor.DisplayName, StringComparison.Ordinal) ||
                !SameFolder(last.Folder, nextAnchor.Folder))
            {
                throw new ArgumentException(
                    "The continuation anchor must identify the final page item.",
                    nameof(nextAnchor));
            }
        }

        private static bool SameFolder(FolderRef left, FolderRef right)
        {
            return string.Equals(left.StoreId, right.StoreId, StringComparison.Ordinal) &&
                string.Equals(left.EntryId, right.EntryId, StringComparison.Ordinal);
        }
    }

    public sealed class OutlookMessagePage
    {
        public OutlookMessagePage(
            IEnumerable<OutlookMessageSummary> items,
            OutlookMessageKeysetAnchor? nextAnchor,
            int totalScopeCount,
            IEnumerable<OutlookScopeFailure> failures)
        {
            var itemCopy = OutlookContractValidation.BoundedCopy(
                items,
                OutlookReadLimits.MaximumPageSize,
                nameof(items),
                rejectNullElements: true);
            ValidateNextAnchor(itemCopy, nextAnchor);
            if (totalScopeCount < 1 ||
                totalScopeCount > OutlookReadLimits.MaximumSearchScopeCount)
            {
                throw new ArgumentOutOfRangeException(nameof(totalScopeCount));
            }

            var failureCopy = OutlookContractValidation.BoundedCopy(
                failures,
                OutlookReadLimits.MaximumSearchScopeCount,
                nameof(failures),
                rejectNullElements: true);
            if (failureCopy.Count >= totalScopeCount)
            {
                throw new ArgumentException(
                    "A successful page requires at least one successful scope.",
                    nameof(failures));
            }

            if (failureCopy.Count > 0 && nextAnchor != null)
            {
                throw new ArgumentException(
                    "A partial cross-scope page cannot have a continuation anchor.",
                    nameof(nextAnchor));
            }

            var failedScopeKeys = new HashSet<string>(StringComparer.Ordinal);
            foreach (var failure in failureCopy)
            {
                if (!failedScopeKeys.Add(failure.Scope.Key))
                {
                    throw new ArgumentException(
                        "A failed scope may appear only once.",
                        nameof(failures));
                }
            }

            Items = itemCopy;
            NextAnchor = nextAnchor;
            TotalScopeCount = totalScopeCount;
            Failures = failureCopy;
            IsPartial = failureCopy.Count > 0;
        }

        public IReadOnlyList<OutlookMessageSummary> Items { get; }

        public OutlookMessageKeysetAnchor? NextAnchor { get; }

        public int TotalScopeCount { get; }

        public IReadOnlyList<OutlookScopeFailure> Failures { get; }

        public bool IsPartial { get; }

        private static void ValidateNextAnchor(
            IReadOnlyList<OutlookMessageSummary> items,
            OutlookMessageKeysetAnchor? nextAnchor)
        {
            if (nextAnchor == null)
            {
                return;
            }

            if (items.Count == 0)
            {
                throw new ArgumentException(
                    "An empty page cannot have a continuation anchor.",
                    nameof(nextAnchor));
            }

            var last = items[items.Count - 1];
            if (last.EffectiveTimestampUtc != nextAnchor.EffectiveTimestampUtc ||
                !string.Equals(last.Item.StoreId, nextAnchor.Item.StoreId, StringComparison.Ordinal) ||
                !string.Equals(last.Item.EntryId, nextAnchor.Item.EntryId, StringComparison.Ordinal) ||
                !string.Equals(last.Item.ItemClass, nextAnchor.Item.ItemClass, StringComparison.Ordinal))
            {
                throw new ArgumentException(
                    "The continuation anchor must identify the final page item.",
                    nameof(nextAnchor));
            }
        }
    }

    public sealed class OutlookAttachmentPage
    {
        public OutlookAttachmentPage(
            IEnumerable<OutlookAttachmentSummary> items,
            OutlookAttachmentKeysetAnchor? nextAnchor)
        {
            var itemCopy = OutlookContractValidation.BoundedCopy(
                items,
                OutlookReadLimits.MaximumPageSize,
                nameof(items),
                rejectNullElements: true);
            ValidateNextAnchor(itemCopy, nextAnchor);
            Items = itemCopy;
            NextAnchor = nextAnchor;
        }

        public IReadOnlyList<OutlookAttachmentSummary> Items { get; }

        public OutlookAttachmentKeysetAnchor? NextAnchor { get; }

        private static void ValidateNextAnchor(
            IReadOnlyList<OutlookAttachmentSummary> items,
            OutlookAttachmentKeysetAnchor? nextAnchor)
        {
            if (nextAnchor == null)
            {
                return;
            }

            if (items.Count == 0)
            {
                throw new ArgumentException(
                    "An empty page cannot have a continuation anchor.",
                    nameof(nextAnchor));
            }

            var last = items[items.Count - 1].Attachment;
            if (last.AttachmentIndex != nextAnchor.AttachmentIndex ||
                !string.Equals(
                    last.MetadataFingerprint,
                    nextAnchor.MetadataFingerprint,
                    StringComparison.Ordinal))
            {
                throw new ArgumentException(
                    "The continuation anchor must identify the final page item.",
                    nameof(nextAnchor));
            }
        }
    }
}
