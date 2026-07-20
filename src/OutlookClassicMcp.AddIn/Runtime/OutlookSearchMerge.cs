using System;
using System.Collections.Generic;
using OutlookClassicMcp.Core.Outlook;

namespace OutlookClassicMcp.AddIn.Runtime
{
    internal sealed class OutlookSearchMerge
    {
        private readonly List<OutlookMessageSummary> _candidates;
        private readonly List<OutlookScopeFailure> _failures =
            new List<OutlookScopeFailure>();
        private readonly int _maximumCount;
        private readonly int _pageSize;
        private readonly int _totalScopeCount;
        private OutlookGatewayException? _firstFailure;
        private int _successfulScopeCount;

        public OutlookSearchMerge(int pageSize, int totalScopeCount)
        {
            if (pageSize < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(pageSize));
            }

            if (totalScopeCount < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(totalScopeCount));
            }

            _pageSize = pageSize;
            _totalScopeCount = totalScopeCount;
            _maximumCount = checked(pageSize + 1);
            _candidates = new List<OutlookMessageSummary>(_maximumCount);
        }

        public static List<OutlookSearchScope> CanonicalizeScopes(
            IReadOnlyList<OutlookSearchScope> scopes)
        {
            if (scopes == null)
            {
                throw new ArgumentNullException(nameof(scopes));
            }

            var sorted = new List<OutlookSearchScope>(scopes);
            sorted.Sort(CompareScopes);
            return sorted;
        }

        public void AddSuccess(IEnumerable<OutlookMessageSummary> items)
        {
            if (items == null)
            {
                throw new ArgumentNullException(nameof(items));
            }

            _successfulScopeCount++;
            foreach (var item in items)
            {
                var candidate = item ?? throw new ArgumentException(
                    "A successful scope returned a null message.",
                    nameof(items));
                if (ContainsExactItem(candidate.Item))
                {
                    continue;
                }

                OutlookReadProjection.InsertBounded(
                    _candidates,
                    candidate,
                    _maximumCount,
                    OutlookReadProjection.CompareMessages);
                OutlookComMetrics.ObserveMaterializedItems(_candidates.Count);
            }
        }

        public void AddFailure(OutlookSearchScope scope, Exception exception)
        {
            if (scope == null)
            {
                throw new ArgumentNullException(nameof(scope));
            }

            var mapped = OutlookErrorMapper.Map(
                exception ?? throw new ArgumentNullException(nameof(exception)));
            if (!IsScopedFailure(mapped.Failure))
            {
                throw mapped;
            }

            _firstFailure = _firstFailure ?? mapped;
            _failures.Add(new OutlookScopeFailure(scope, mapped.Failure));
        }

        public OutlookMessagePage Complete()
        {
            if (_successfulScopeCount == 0)
            {
                throw _firstFailure ??
                    new OutlookGatewayException(OutlookGatewayFailure.Internal);
            }

            return OutlookReadProjection.BuildMessagePage(
                _candidates,
                _pageSize,
                _totalScopeCount,
                _failures);
        }

        private bool ContainsExactItem(ItemRef item)
        {
            foreach (var candidate in _candidates)
            {
                if (string.Equals(
                        candidate.Item.StoreId,
                        item.StoreId,
                        StringComparison.Ordinal) &&
                    string.Equals(
                        candidate.Item.EntryId,
                        item.EntryId,
                        StringComparison.Ordinal) &&
                    string.Equals(
                        candidate.Item.ItemClass,
                        item.ItemClass,
                        StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        private static int CompareScopes(OutlookSearchScope left, OutlookSearchScope right)
        {
            var store = string.Compare(
                left.Mailbox.StoreId,
                right.Mailbox.StoreId,
                StringComparison.Ordinal);
            return store != 0
                ? store
                : string.Compare(
                    left.Folder?.EntryId ?? string.Empty,
                    right.Folder?.EntryId ?? string.Empty,
                    StringComparison.Ordinal);
        }

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
