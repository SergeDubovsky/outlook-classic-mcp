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
        private int _listMailboxesCallCount;
        private int _listFoldersCallCount;
        private int _listMessagesCallCount;
        private int _searchMessagesCallCount;
        private int _getMessageCallCount;
        private int _getConversationCallCount;
        private int _listAttachmentsCallCount;

        public FakeOutlookGateway()
            : this(_ => Task.FromResult(ProbeTestData.CreateSnapshot()))
        {
        }

        public FakeOutlookGateway(Func<CancellationToken, Task<OutlookProbeSnapshot>> handler)
        {
            _handler = handler ?? throw new ArgumentNullException(nameof(handler));
        }

        public int CallCount => Volatile.Read(ref _callCount);

        public int ListMailboxesCallCount => Volatile.Read(ref _listMailboxesCallCount);

        public int ListFoldersCallCount => Volatile.Read(ref _listFoldersCallCount);

        public int ListMessagesCallCount => Volatile.Read(ref _listMessagesCallCount);

        public int SearchMessagesCallCount => Volatile.Read(ref _searchMessagesCallCount);

        public int GetMessageCallCount => Volatile.Read(ref _getMessageCallCount);

        public int GetConversationCallCount => Volatile.Read(ref _getConversationCallCount);

        public int ListAttachmentsCallCount => Volatile.Read(ref _listAttachmentsCallCount);

        public int ReadCallCount =>
            ListMailboxesCallCount +
            ListFoldersCallCount +
            ListMessagesCallCount +
            SearchMessagesCallCount +
            GetMessageCallCount +
            GetConversationCallCount +
            ListAttachmentsCallCount;

        public CancellationToken LastCancellationToken { get; private set; }

        public Func<OutlookListMailboxesRequest, CancellationToken, Task<OutlookMailboxPage>>?
            ListMailboxesHandler { get; set; }

        public Func<OutlookListFoldersRequest, CancellationToken, Task<OutlookFolderPage>>?
            ListFoldersHandler { get; set; }

        public Func<OutlookListMessagesRequest, CancellationToken, Task<OutlookMessagePage>>?
            ListMessagesHandler { get; set; }

        public Func<OutlookSearchMessagesRequest, CancellationToken, Task<OutlookMessagePage>>?
            SearchMessagesHandler { get; set; }

        public Func<OutlookGetMessageRequest, CancellationToken, Task<OutlookMessageDetail>>?
            GetMessageHandler { get; set; }

        public Func<OutlookGetConversationRequest, CancellationToken, Task<OutlookMessagePage>>?
            GetConversationHandler { get; set; }

        public Func<OutlookListAttachmentsRequest, CancellationToken, Task<OutlookAttachmentPage>>?
            ListAttachmentsHandler { get; set; }

        public OutlookListMailboxesRequest? LastListMailboxesRequest { get; private set; }

        public OutlookListFoldersRequest? LastListFoldersRequest { get; private set; }

        public OutlookListMessagesRequest? LastListMessagesRequest { get; private set; }

        public OutlookSearchMessagesRequest? LastSearchMessagesRequest { get; private set; }

        public OutlookGetMessageRequest? LastGetMessageRequest { get; private set; }

        public OutlookGetConversationRequest? LastGetConversationRequest { get; private set; }

        public OutlookListAttachmentsRequest? LastListAttachmentsRequest { get; private set; }

        public Task<OutlookProbeSnapshot> ProbeAsync(CancellationToken cancellationToken)
        {
            Interlocked.Increment(ref _callCount);
            LastCancellationToken = cancellationToken;
            return _handler(cancellationToken);
        }

        public Task<OutlookMailboxPage> ListMailboxesAsync(
            OutlookListMailboxesRequest request,
            CancellationToken cancellationToken)
        {
            Interlocked.Increment(ref _listMailboxesCallCount);
            LastListMailboxesRequest = request;
            LastCancellationToken = cancellationToken;
            return RequireHandler(ListMailboxesHandler)(request, cancellationToken);
        }

        public Task<OutlookFolderPage> ListFoldersAsync(
            OutlookListFoldersRequest request,
            CancellationToken cancellationToken)
        {
            Interlocked.Increment(ref _listFoldersCallCount);
            LastListFoldersRequest = request;
            LastCancellationToken = cancellationToken;
            return RequireHandler(ListFoldersHandler)(request, cancellationToken);
        }

        public Task<OutlookMessagePage> ListMessagesAsync(
            OutlookListMessagesRequest request,
            CancellationToken cancellationToken)
        {
            Interlocked.Increment(ref _listMessagesCallCount);
            LastListMessagesRequest = request;
            LastCancellationToken = cancellationToken;
            return RequireHandler(ListMessagesHandler)(request, cancellationToken);
        }

        public Task<OutlookMessagePage> SearchMessagesAsync(
            OutlookSearchMessagesRequest request,
            CancellationToken cancellationToken)
        {
            Interlocked.Increment(ref _searchMessagesCallCount);
            LastSearchMessagesRequest = request;
            LastCancellationToken = cancellationToken;
            return RequireHandler(SearchMessagesHandler)(request, cancellationToken);
        }

        public Task<OutlookMessageDetail> GetMessageAsync(
            OutlookGetMessageRequest request,
            CancellationToken cancellationToken)
        {
            Interlocked.Increment(ref _getMessageCallCount);
            LastGetMessageRequest = request;
            LastCancellationToken = cancellationToken;
            return RequireHandler(GetMessageHandler)(request, cancellationToken);
        }

        public Task<OutlookMessagePage> GetConversationAsync(
            OutlookGetConversationRequest request,
            CancellationToken cancellationToken)
        {
            Interlocked.Increment(ref _getConversationCallCount);
            LastGetConversationRequest = request;
            LastCancellationToken = cancellationToken;
            return RequireHandler(GetConversationHandler)(request, cancellationToken);
        }

        public Task<OutlookAttachmentPage> ListAttachmentsAsync(
            OutlookListAttachmentsRequest request,
            CancellationToken cancellationToken)
        {
            Interlocked.Increment(ref _listAttachmentsCallCount);
            LastListAttachmentsRequest = request;
            LastCancellationToken = cancellationToken;
            return RequireHandler(ListAttachmentsHandler)(request, cancellationToken);
        }

        private static THandler RequireHandler<THandler>(THandler? handler)
            where THandler : class
        {
            return handler ?? throw new InvalidOperationException(
                "The fake gateway handler was not configured for this operation.");
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
