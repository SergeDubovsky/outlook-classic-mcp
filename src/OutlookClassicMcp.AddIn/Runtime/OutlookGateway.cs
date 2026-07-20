using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using OutlookClassicMcp.Core.Outlook;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookClassicMcp.AddIn.Runtime
{
    internal sealed partial class OutlookGateway : IOutlookGateway
    {
        private const int MapiNotFound = unchecked((int)0x8004010F);
        private const int RpcCallRejected = unchecked((int)0x80010001);
        private const int RpcDisconnected = unchecked((int)0x80010108);
        private const int RpcServerCallRetryLater = unchecked((int)0x8001010A);
        private const int ObjectNotConnected = unchecked((int)0x800401FD);
        private readonly OutlookStaDispatcher _dispatcher;

        public OutlookGateway(OutlookStaDispatcher dispatcher)
        {
            _dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
        }

        public async Task<OutlookProbeSnapshot> ProbeAsync(CancellationToken cancellationToken)
        {
            try
            {
                return await _dispatcher
                    .InvokeWithContextAsync<OutlookRuntimeContext, OutlookProbeSnapshot>(
                        ProbeOnOutlookThread,
                        cancellationToken)
                    .ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (Exception exception)
            {
                throw OutlookErrorMapper.Map(exception);
            }
        }

        private static OutlookProbeSnapshot ProbeOnOutlookThread(OutlookRuntimeContext context)
        {
            var executedThread = context.AssertAndCaptureCurrentThread();
            var capturedThread = context.CapturedThread;
            var threadProof = new OutlookDispatcherThreadProof(
                capturedThread.ManagedThreadId,
                capturedThread.NativeThreadId,
                executedThread.ManagedThreadId,
                executedThread.NativeThreadId,
                executedThread.ApartmentState == ApartmentState.STA);

            var application = context.Application;
            var metadataIncomplete = false;
            var version = BoundRequiredText(
                application.Version,
                OutlookProbeSnapshot.MaximumVersionLength,
                ref metadataIncomplete);

            // Application and Application.Session are host-owned and are never released here.
            var session = application.Session;
            if (session == null)
            {
                throw new OutlookGatewayException(OutlookGatewayFailure.NotReady);
            }

            var profileName = BoundRequiredText(
                session.CurrentProfileName,
                OutlookProbeSnapshot.MaximumProfileNameLength,
                ref metadataIncomplete);

            Outlook.Stores? stores = null;
            try
            {
                stores = session.Stores;
                if (stores == null)
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.NotReady);
                }
                OutlookComMetrics.RecordAcquired();

                var configuredStoreCount = stores.Count;
                var returnedStoreCount = Math.Min(
                    configuredStoreCount,
                    OutlookProbeSnapshot.MaximumStoreCount);
                var results = new List<OutlookStoreProbe>(returnedStoreCount);

                for (var index = 1; index <= returnedStoreCount; index++)
                {
                    Outlook.Store? store = null;
                    try
                    {
                        store = stores[index];
                        if (store == null)
                        {
                            metadataIncomplete = true;
                            continue;
                        }
                        OutlookComMetrics.RecordAcquired();

                        var result = ProbeStore(store, ref metadataIncomplete);
                        if (result != null)
                        {
                            results.Add(result);
                            OutlookComMetrics.ObserveMaterializedItems(results.Count);
                        }
                    }
                    finally
                    {
                        if (store != null)
                        {
                            Marshal.ReleaseComObject(store);
                            OutlookComMetrics.RecordReleased();
                        }
                    }
                }

                var warnings = new List<OutlookProbeWarning>(3)
                {
                    OutlookProbeWarning.ArchiveNotExposedByOutlookObjectModel,
                };
                if (metadataIncomplete)
                {
                    warnings.Add(OutlookProbeWarning.StoreMetadataIncomplete);
                }

                if (configuredStoreCount > OutlookProbeSnapshot.MaximumStoreCount)
                {
                    warnings.Add(OutlookProbeWarning.StoreLimitReached);
                }

                return new OutlookProbeSnapshot(
                    version,
                    IntPtr.Size * 8,
                    profileName,
                    threadProof,
                    configuredStoreCount,
                    results,
                    warnings);
            }
            finally
            {
                if (stores != null)
                {
                    Marshal.ReleaseComObject(stores);
                    OutlookComMetrics.RecordReleased();
                }
            }
        }

        private static OutlookStoreProbe? ProbeStore(
            Outlook.Store store,
            ref bool metadataIncomplete)
        {
            var displayName = ReadDisplayName(store, ref metadataIncomplete);
            if (displayName == null)
            {
                return null;
            }

            var storeType = ReadStoreType(store, ref metadataIncomplete);
            var isDataFileStore = ReadIsDataFileStore(store, ref metadataIncomplete);
            var isCachedExchange = ReadIsCachedExchange(store, ref metadataIncomplete);
            var isOpen = ReadIsOpen(store, ref metadataIncomplete);

            var capabilities = new OutlookStoreCapabilities(
                storeType != OutlookStoreType.NonExchange &&
                    storeType != OutlookStoreType.Unknown,
                isDataFileStore,
                isCachedExchange);

            StandardFolderAvailability standardFolders;
            if (!isOpen)
            {
                standardFolders = UnknownFolderAvailability();
            }
            else
            {
                standardFolders = new StandardFolderAvailability(
                    ProbeDefaultFolder(
                        store,
                        Outlook.OlDefaultFolders.olFolderInbox,
                        ref metadataIncomplete),
                    ProbeDefaultFolder(
                        store,
                        Outlook.OlDefaultFolders.olFolderDrafts,
                        ref metadataIncomplete),
                    ProbeDefaultFolder(
                        store,
                        Outlook.OlDefaultFolders.olFolderSentMail,
                        ref metadataIncomplete),
                    ProbeDefaultFolder(
                        store,
                        Outlook.OlDefaultFolders.olFolderDeletedItems,
                        ref metadataIncomplete),
                    OutlookFolderAvailability.Unknown);
            }

            return new OutlookStoreProbe(
                displayName,
                storeType,
                capabilities,
                standardFolders);
        }

        private static string? ReadDisplayName(
            Outlook.Store store,
            ref bool metadataIncomplete)
        {
            try
            {
                var displayName = store.DisplayName;
                if (string.IsNullOrWhiteSpace(displayName))
                {
                    metadataIncomplete = true;
                    return null;
                }

                return BoundRequiredText(
                    displayName,
                    OutlookStoreProbe.MaximumDisplayNameLength,
                    ref metadataIncomplete);
            }
            catch (COMException exception) when (!IsFatalProviderFailure(exception))
            {
                metadataIncomplete = true;
                return null;
            }
        }

        private static OutlookStoreType ReadStoreType(
            Outlook.Store store,
            ref bool metadataIncomplete)
        {
            try
            {
                return MapStoreType(store.ExchangeStoreType);
            }
            catch (COMException exception) when (!IsFatalProviderFailure(exception))
            {
                metadataIncomplete = true;
                return OutlookStoreType.Unknown;
            }
        }

        private static bool ReadIsDataFileStore(
            Outlook.Store store,
            ref bool metadataIncomplete)
        {
            try
            {
                return store.IsDataFileStore;
            }
            catch (COMException exception) when (!IsFatalProviderFailure(exception))
            {
                metadataIncomplete = true;
                return false;
            }
        }

        private static bool ReadIsCachedExchange(
            Outlook.Store store,
            ref bool metadataIncomplete)
        {
            try
            {
                return store.IsCachedExchange;
            }
            catch (COMException exception) when (!IsFatalProviderFailure(exception))
            {
                metadataIncomplete = true;
                return false;
            }
        }

        private static bool ReadIsOpen(
            Outlook.Store store,
            ref bool metadataIncomplete)
        {
            try
            {
                return store.IsOpen;
            }
            catch (COMException exception) when (!IsFatalProviderFailure(exception))
            {
                metadataIncomplete = true;
                return false;
            }
        }

        private static OutlookFolderAvailability ProbeDefaultFolder(
            Outlook.Store store,
            Outlook.OlDefaultFolders folderKind,
            ref bool metadataIncomplete)
        {
            Outlook.MAPIFolder? folder = null;
            try
            {
                folder = store.GetDefaultFolder(folderKind);
                if (folder != null)
                {
                    OutlookComMetrics.RecordAcquired();
                }
                return folder == null
                    ? OutlookFolderAvailability.Missing
                    : OutlookFolderAvailability.Available;
            }
            catch (COMException exception) when (exception.ErrorCode == MapiNotFound)
            {
                return OutlookFolderAvailability.Missing;
            }
            catch (COMException exception) when (!IsFatalProviderFailure(exception))
            {
                metadataIncomplete = true;
                return OutlookFolderAvailability.Unknown;
            }
            finally
            {
                if (folder != null)
                {
                    Marshal.ReleaseComObject(folder);
                    OutlookComMetrics.RecordReleased();
                }
            }
        }

        private static StandardFolderAvailability UnknownFolderAvailability()
        {
            return new StandardFolderAvailability(
                OutlookFolderAvailability.Unknown,
                OutlookFolderAvailability.Unknown,
                OutlookFolderAvailability.Unknown,
                OutlookFolderAvailability.Unknown,
                OutlookFolderAvailability.Unknown);
        }

        private static OutlookStoreType MapStoreType(Outlook.OlExchangeStoreType storeType)
        {
            switch (storeType)
            {
                case Outlook.OlExchangeStoreType.olPrimaryExchangeMailbox:
                    return OutlookStoreType.PrimaryExchangeMailbox;
                case Outlook.OlExchangeStoreType.olExchangeMailbox:
                    return OutlookStoreType.ExchangeMailbox;
                case Outlook.OlExchangeStoreType.olExchangePublicFolder:
                    return OutlookStoreType.ExchangePublicFolder;
                case Outlook.OlExchangeStoreType.olAdditionalExchangeMailbox:
                    return OutlookStoreType.AdditionalExchangeMailbox;
                case Outlook.OlExchangeStoreType.olNotExchange:
                    return OutlookStoreType.NonExchange;
                default:
                    return OutlookStoreType.Unknown;
            }
        }

        private static bool IsFatalProviderFailure(COMException exception)
        {
            switch (exception.ErrorCode)
            {
                case RpcCallRejected:
                case RpcServerCallRetryLater:
                case RpcDisconnected:
                case ObjectNotConnected:
                    return true;
                default:
                    return false;
            }
        }

        private static string BoundRequiredText(
            string value,
            int maximumLength,
            ref bool metadataIncomplete)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new InvalidOperationException("Outlook returned incomplete host metadata.");
            }

            if (value.Length <= maximumLength)
            {
                return value;
            }

            metadataIncomplete = true;
            return value.Substring(0, maximumLength);
        }
    }
}
