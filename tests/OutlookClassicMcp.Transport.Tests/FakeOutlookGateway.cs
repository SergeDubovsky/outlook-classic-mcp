using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OutlookClassicMcp.Core.Outlook;

namespace OutlookClassicMcp.Transport.Tests
{
    internal sealed class FakeOutlookGateway : IOutlookGateway
    {
        private readonly Func<CancellationToken, Task<OutlookProbeSnapshot>> _handler;
        private int _callCount;

        public FakeOutlookGateway()
            : this(_ => Task.FromResult(ProbeTestData.CreateSnapshot()))
        {
        }

        public FakeOutlookGateway(Func<CancellationToken, Task<OutlookProbeSnapshot>> handler)
        {
            _handler = handler ?? throw new ArgumentNullException(nameof(handler));
        }

        public int CallCount => Volatile.Read(ref _callCount);

        public CancellationToken LastCancellationToken { get; private set; }

        public Task<OutlookProbeSnapshot> ProbeAsync(CancellationToken cancellationToken)
        {
            Interlocked.Increment(ref _callCount);
            LastCancellationToken = cancellationToken;
            return _handler(cancellationToken);
        }
    }

    internal static class ProbeTestData
    {
        public static OutlookProbeSnapshot CreateSnapshot(
            IEnumerable<OutlookStoreProbe>? stores = null,
            int configuredStoreCount = 2,
            IEnumerable<OutlookProbeWarning>? warnings = null)
        {
            var storeValues = stores ?? new[]
            {
                CreateStore(
                    "Primary mailbox",
                    OutlookStoreType.PrimaryExchangeMailbox,
                    isExchangeStore: true,
                    isDataFileStore: false,
                    isCachedExchange: true,
                    OutlookFolderAvailability.Available,
                    OutlookFolderAvailability.Available,
                    OutlookFolderAvailability.Available,
                    OutlookFolderAvailability.Available),
                CreateStore(
                    "Local data file",
                    OutlookStoreType.NonExchange,
                    isExchangeStore: false,
                    isDataFileStore: true,
                    isCachedExchange: false,
                    OutlookFolderAvailability.Available,
                    OutlookFolderAvailability.Missing,
                    OutlookFolderAvailability.Missing,
                    OutlookFolderAvailability.Available),
            };
            var warningValues = warnings ?? new[]
            {
                OutlookProbeWarning.ArchiveNotExposedByOutlookObjectModel,
            };

            return new OutlookProbeSnapshot(
                "16.0.12345.10000",
                64,
                "Test Profile",
                new OutlookDispatcherThreadProof(17, 23, 17, 23, executedOnSta: true),
                configuredStoreCount,
                storeValues,
                warningValues);
        }

        public static OutlookStoreProbe CreateStore(
            string displayName,
            OutlookStoreType storeType,
            bool isExchangeStore,
            bool isDataFileStore,
            bool isCachedExchange,
            OutlookFolderAvailability inbox,
            OutlookFolderAvailability drafts,
            OutlookFolderAvailability sent,
            OutlookFolderAvailability deleted)
        {
            return new OutlookStoreProbe(
                displayName,
                storeType,
                new OutlookStoreCapabilities(
                    isExchangeStore,
                    isDataFileStore,
                    isCachedExchange),
                new StandardFolderAvailability(
                    inbox,
                    drafts,
                    sent,
                    deleted,
                    OutlookFolderAvailability.Unknown));
        }
    }
}
