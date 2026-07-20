using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using OutlookClassicMcp.Core.Outlook;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookClassicMcp.AddIn.Runtime
{
    internal sealed class OutlookSearchScopeState
    {
        public OutlookSearchScopeState(
            OutlookSearchScope scope,
            OutlookMessageSearchFilter filter,
            OutlookMessageKeysetAnchor? anchor,
            int pageSize)
        {
            Scope = scope;
            Filter = filter;
            Anchor = anchor;
            PageSize = pageSize;
        }

        public OutlookSearchScope Scope { get; }

        public OutlookMessageSearchFilter Filter { get; }

        public OutlookMessageKeysetAnchor? Anchor { get; }

        public int PageSize { get; }
    }

    internal sealed class OutlookSearchAnchorState
    {
        public OutlookSearchAnchorState(
            OutlookSearchMessagesRequest request,
            IReadOnlyList<OutlookSearchScope> scopes)
        {
            Request = request ?? throw new ArgumentNullException(nameof(request));
            Scopes = scopes ?? throw new ArgumentNullException(nameof(scopes));
        }

        public OutlookSearchMessagesRequest Request { get; }

        public IReadOnlyList<OutlookSearchScope> Scopes { get; }
    }

    internal static class OutlookReadOperations
    {
        private const int MaximumChildFoldersExamined = 1024;
        private const int MaximumConversationDepth = 64;
        private const int MaximumConversationNodesExamined = 1024;
        private const int MaximumMailboxesExamined = 64;
        private const int MaximumMessagesExamined = 4096;
        private const int MaximumMessageTimestampTieGroup = 1024;
        private const int MaximumRecipientsExamined = 1024;
        private const int MapiNotSupported = unchecked((int)0x80040102);
        private const string AttachmentMimeTagSchema =
            "https://schemas.microsoft.com/mapi/proptag/0x370E001F";

        public static OutlookMailboxPage ListMailboxes(
            OutlookRuntimeContext context,
            OutlookListMailboxesRequest request,
            CancellationToken cancellationToken)
        {
            context.AssertAndCaptureCurrentThread();
            cancellationToken.ThrowIfCancellationRequested();
            var session = RequireSession(context);
            if (request.Anchor != null)
            {
                ValidateMailboxAnchor(session, request.Anchor);
            }

            Outlook.Stores? stores = null;
            try
            {
                stores = session.Stores;
                if (stores == null)
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.NotReady);
                }
                OutlookComMetrics.RecordAcquired();

                if (stores.Count > MaximumMailboxesExamined)
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.Timeout);
                }

                var candidates = new List<OutlookMailboxSummary>(request.PageSize + 1);
                for (var index = 1; index <= stores.Count; index++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    Outlook.Store? store = null;
                    try
                    {
                        store = stores[index];
                        if (store == null)
                        {
                            continue;
                        }
                        OutlookComMetrics.RecordAcquired();

                        var summary = ProjectMailbox(store);
                        if (request.Anchor == null ||
                            OutlookReadProjection.CompareMailboxToAnchor(summary, request.Anchor) > 0)
                        {
                            OutlookReadProjection.InsertBounded(
                                candidates,
                                summary,
                                request.PageSize + 1,
                                OutlookReadProjection.CompareMailboxes);
                            OutlookComMetrics.ObserveMaterializedItems(candidates.Count);
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

                return BuildMailboxPage(candidates, request.PageSize);
            }
            catch (COMException exception) when (exception.ErrorCode == MapiNotSupported)
            {
                throw new OutlookGatewayException(OutlookGatewayFailure.UnsupportedStore);
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

        public static OutlookFolderPage ListFolders(
            OutlookRuntimeContext context,
            OutlookListFoldersRequest request,
            CancellationToken cancellationToken)
        {
            context.AssertAndCaptureCurrentThread();
            cancellationToken.ThrowIfCancellationRequested();
            var session = RequireSession(context);

            Outlook.Store? store = null;
            Outlook.MAPIFolder? container = null;
            Outlook.Folders? folders = null;
            try
            {
                store = GetStore(session, request.Mailbox.StoreId);
                container = request.ParentFolder == null
                    ? GetRootFolder(store)
                    : GetFolder(session, request.ParentFolder);

                if (request.Anchor != null)
                {
                    ValidateFolderAnchor(session, container, request.Anchor);
                }

                folders = container.Folders;
                if (folders == null)
                {
                    return new OutlookFolderPage(
                        Array.Empty<OutlookFolderSummary>(),
                        nextAnchor: null);
                }
                OutlookComMetrics.RecordAcquired();

                if (folders.Count > MaximumChildFoldersExamined)
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.Timeout);
                }

                var candidates = new List<OutlookFolderSummary>(request.PageSize + 1);
                for (var index = 1; index <= folders.Count; index++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    Outlook.MAPIFolder? child = null;
                    Outlook.Folders? childFolders = null;
                    try
                    {
                        child = folders[index];
                        if (child == null)
                        {
                            continue;
                        }
                        OutlookComMetrics.RecordAcquired();

                        childFolders = child.Folders;
                        if (childFolders != null)
                        {
                            OutlookComMetrics.RecordAcquired();
                        }
                        var summary = new OutlookFolderSummary(
                            new FolderRef(
                                OutlookReadProjection.RequireIdentifier(
                                    child.StoreID,
                                    MailboxRef.MaximumStoreIdLength),
                                OutlookReadProjection.RequireIdentifier(
                                    child.EntryID,
                                    FolderRef.MaximumEntryIdLength)),
                            request.ParentFolder,
                            OutlookReadProjection.BoundDisplay(
                                child.Name,
                                OutlookFolderSummary.MaximumDisplayNameLength),
                            childFolders != null && childFolders.Count > 0);

                        if (request.Anchor == null ||
                            OutlookReadProjection.CompareFolderToAnchor(summary, request.Anchor) > 0)
                        {
                            OutlookReadProjection.InsertBounded(
                                candidates,
                                summary,
                                request.PageSize + 1,
                                OutlookReadProjection.CompareFolders);
                            OutlookComMetrics.ObserveMaterializedItems(candidates.Count);
                        }
                    }
                    finally
                    {
                        if (childFolders != null)
                        {
                            Marshal.ReleaseComObject(childFolders);
                            OutlookComMetrics.RecordReleased();
                        }

                        if (child != null)
                        {
                            Marshal.ReleaseComObject(child);
                            OutlookComMetrics.RecordReleased();
                        }
                    }
                }

                return BuildFolderPage(candidates, request.PageSize);
            }
            catch (COMException exception) when (exception.ErrorCode == MapiNotSupported)
            {
                throw new OutlookGatewayException(OutlookGatewayFailure.UnsupportedStore);
            }
            finally
            {
                if (folders != null)
                {
                    Marshal.ReleaseComObject(folders);
                    OutlookComMetrics.RecordReleased();
                }

                if (container != null)
                {
                    Marshal.ReleaseComObject(container);
                    OutlookComMetrics.RecordReleased();
                }

                if (store != null)
                {
                    Marshal.ReleaseComObject(store);
                    OutlookComMetrics.RecordReleased();
                }
            }
        }

        public static OutlookMessagePage ListMessages(
            OutlookRuntimeContext context,
            OutlookListMessagesRequest request,
            CancellationToken cancellationToken)
        {
            context.AssertAndCaptureCurrentThread();
            cancellationToken.ThrowIfCancellationRequested();
            var session = RequireSession(context);

            Outlook.Store? store = null;
            Outlook.MAPIFolder? folder = null;
            try
            {
                store = GetStore(session, request.Folder.StoreId);
                folder = GetFolder(session, request.Folder);
                RejectVisibleSearchFolder(
                    session,
                    store,
                    folder,
                    cancellationToken);
                var timestampKind = ResolveTimestampKind(session, store, folder);
                if (request.Anchor != null)
                {
                    ValidateMessageAnchor(
                        session,
                        request.Anchor,
                        request.Folder,
                        timestampKind,
                        filter: null,
                        expectedConversationId: null);
                }

                var candidates = ReadMessagesFromFolder(
                    folder,
                    request.Folder,
                    OutlookReadFilter.BuildListRestriction(timestampKind, request.Anchor),
                    timestampKind,
                    request.Anchor,
                    filter: null,
                    request.PageSize,
                    cancellationToken);
                return OutlookReadProjection.BuildMessagePage(
                    candidates,
                    request.PageSize,
                    1,
                    Array.Empty<OutlookScopeFailure>());
            }
            finally
            {
                if (folder != null)
                {
                    Marshal.ReleaseComObject(folder);
                    OutlookComMetrics.RecordReleased();
                }

                if (store != null)
                {
                    Marshal.ReleaseComObject(store);
                    OutlookComMetrics.RecordReleased();
                }
            }
        }

        public static bool ValidateSearchAnchor(
            OutlookRuntimeContext context,
            OutlookSearchAnchorState state,
            CancellationToken cancellationToken)
        {
            context.AssertAndCaptureCurrentThread();
            cancellationToken.ThrowIfCancellationRequested();
            var request = state.Request;
            if (request.Anchor == null)
            {
                return true;
            }

            var session = RequireSession(context);
            Outlook.Store? store = null;
            object? rawItem = null;
            Outlook.MAPIFolder? anchorFolder = null;
            try
            {
                store = GetStore(session, request.Anchor.Item.StoreId);
                rawItem = GetItem(session, request.Anchor.Item);
                if (!(rawItem is Outlook.MailItem mail))
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.CursorStale);
                }

                RequireMatchingItemClass(mail, request.Anchor.Item, OutlookGatewayFailure.CursorStale);
                var parent = OutlookReadProjection.ReadParentFolder(mail);
                anchorFolder = GetFolder(session, parent);
                var matchingScope = false;
                foreach (var scope in state.Scopes)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!string.Equals(
                        scope.Mailbox.StoreId,
                        parent.StoreId,
                        StringComparison.Ordinal))
                    {
                        continue;
                    }

                    if (scope.Folder != null)
                    {
                        if (SameFolder(scope.Folder, parent))
                        {
                            matchingScope = true;
                            break;
                        }

                        continue;
                    }

                    Outlook.MAPIFolder? inbox = null;
                    try
                    {
                        inbox = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                        if (inbox != null)
                        {
                            OutlookComMetrics.RecordAcquired();
                        }
                        if (inbox != null && string.Equals(
                            inbox.EntryID,
                            parent.EntryId,
                            StringComparison.Ordinal))
                        {
                            matchingScope = true;
                            break;
                        }
                    }
                    finally
                    {
                        if (inbox != null)
                        {
                            Marshal.ReleaseComObject(inbox);
                            OutlookComMetrics.RecordReleased();
                        }
                    }
                }

                if (!matchingScope || !OutlookReadFilter.Matches(mail, request.Filter))
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.CursorStale);
                }

                var timestampKind = ResolveTimestampKind(session, store, anchorFolder);
                var projected = OutlookReadProjection.ProjectMessage(
                    mail,
                    parent,
                    timestampKind);
                if (projected.EffectiveTimestampUtc != request.Anchor.EffectiveTimestampUtc)
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.CursorStale);
                }

                return true;
            }
            catch (OutlookGatewayException exception)
                when (exception.Failure == OutlookGatewayFailure.ItemNotFound)
            {
                throw new OutlookGatewayException(OutlookGatewayFailure.CursorStale);
            }
            catch (COMException exception) when (exception.ErrorCode == MapiNotSupported)
            {
                throw new OutlookGatewayException(OutlookGatewayFailure.UnsupportedStore);
            }
            finally
            {
                if (anchorFolder != null)
                {
                    Marshal.ReleaseComObject(anchorFolder);
                    OutlookComMetrics.RecordReleased();
                }

                if (rawItem != null && Marshal.IsComObject(rawItem))
                {
                    Marshal.ReleaseComObject(rawItem);
                    OutlookComMetrics.RecordReleased();
                }

                if (store != null)
                {
                    Marshal.ReleaseComObject(store);
                    OutlookComMetrics.RecordReleased();
                }
            }
        }

        public static List<OutlookMessageSummary> SearchScope(
            OutlookRuntimeContext context,
            OutlookSearchScopeState state,
            CancellationToken cancellationToken)
        {
            context.AssertAndCaptureCurrentThread();
            cancellationToken.ThrowIfCancellationRequested();
            var session = RequireSession(context);

            Outlook.Store? store = null;
            Outlook.MAPIFolder? folder = null;
            try
            {
                store = GetStore(session, state.Scope.Mailbox.StoreId);
                folder = state.Scope.Folder == null
                    ? GetDefaultFolder(store, Outlook.OlDefaultFolders.olFolderInbox)
                    : GetFolder(session, state.Scope.Folder);
                if (state.Scope.Folder != null)
                {
                    RejectVisibleSearchFolder(
                        session,
                        store,
                        folder,
                        cancellationToken);
                }
                var folderRef = state.Scope.Folder ?? new FolderRef(
                    OutlookReadProjection.RequireIdentifier(
                        folder.StoreID,
                        MailboxRef.MaximumStoreIdLength),
                    OutlookReadProjection.RequireIdentifier(
                        folder.EntryID,
                        FolderRef.MaximumEntryIdLength));
                var timestampKind = ResolveTimestampKind(session, store, folder);

                return ReadMessagesFromFolder(
                    folder,
                    folderRef,
                    OutlookReadFilter.BuildSearchRestriction(
                        state.Filter,
                        state.Anchor,
                        timestampKind),
                    timestampKind,
                    state.Anchor,
                    state.Filter,
                    state.PageSize,
                    cancellationToken);
            }
            catch (COMException exception) when (exception.ErrorCode == MapiNotSupported)
            {
                throw new OutlookGatewayException(OutlookGatewayFailure.UnsupportedStore);
            }
            finally
            {
                if (folder != null)
                {
                    Marshal.ReleaseComObject(folder);
                    OutlookComMetrics.RecordReleased();
                }

                if (store != null)
                {
                    Marshal.ReleaseComObject(store);
                    OutlookComMetrics.RecordReleased();
                }
            }
        }

        public static OutlookMessageDetail GetMessage(
            OutlookRuntimeContext context,
            OutlookGetMessageRequest request,
            CancellationToken cancellationToken)
        {
            context.AssertAndCaptureCurrentThread();
            cancellationToken.ThrowIfCancellationRequested();
            var session = RequireSession(context);

            Outlook.Store? store = null;
            object? rawItem = null;
            Outlook.Recipients? recipients = null;
            try
            {
                store = GetStore(session, request.Item.StoreId);
                rawItem = GetItem(session, request.Item);
                if (!(rawItem is Outlook.MailItem mail))
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.UnsupportedItemType);
                }

                RequireMatchingItemClass(mail, request.Item, OutlookGatewayFailure.ItemMovedOrDeleted);
                var parent = OutlookReadProjection.ReadParentFolder(mail);
                if (!string.Equals(parent.StoreId, request.Item.StoreId, StringComparison.Ordinal))
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.ItemMovedOrDeleted);
                }

                var summary = OutlookReadProjection.ProjectMessage(
                    mail,
                    parent,
                    OutlookMessageTimestampKind.Automatic);
                var toRecipients = new List<OutlookMessageAddress>();
                var ccRecipients = new List<OutlookMessageAddress>();
                var bccRecipients = new List<OutlookMessageAddress>();
                var totalTo = 0;
                var totalCc = 0;
                var totalBcc = 0;

                try
                {
                    recipients = mail.Recipients;
                    if (recipients != null)
                    {
                        OutlookComMetrics.RecordAcquired();
                    }
                }
                catch (COMException exception)
                    when (OutlookErrorMapper.IsObjectModelGuardDenied(exception))
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.ObjectModelGuard);
                }

                if (recipients != null)
                {
                    if (recipients.Count > MaximumRecipientsExamined)
                    {
                        throw new OutlookGatewayException(OutlookGatewayFailure.Timeout);
                    }

                    for (var index = 1; index <= recipients.Count; index++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        Outlook.Recipient? recipient = null;
                        try
                        {
                            recipient = recipients[index];
                            if (recipient == null)
                            {
                                continue;
                            }
                            OutlookComMetrics.RecordAcquired();

                            List<OutlookMessageAddress>? target;
                            switch ((Outlook.OlMailRecipientType)recipient.Type)
                            {
                                case Outlook.OlMailRecipientType.olTo:
                                    totalTo++;
                                    target = toRecipients;
                                    break;
                                case Outlook.OlMailRecipientType.olCC:
                                    totalCc++;
                                    target = ccRecipients;
                                    break;
                                case Outlook.OlMailRecipientType.olBCC:
                                    totalBcc++;
                                    target = bccRecipients;
                                    break;
                                default:
                                    target = null;
                                    break;
                            }

                            if (target == null || target.Count == OutlookMessageDetail.MaximumRecipientCount)
                            {
                                continue;
                            }

                            var name = OutlookReadProjection.BoundOptionalText(
                                recipient.Name,
                                OutlookMessageAddress.MaximumValueLength);
                            var address = OutlookReadProjection.BoundOptionalText(
                                recipient.Address,
                                OutlookMessageAddress.MaximumValueLength);
                            if (name != null || address != null)
                            {
                                target.Add(new OutlookMessageAddress(name, address));
                            }
                        }
                        catch (COMException exception)
                            when (OutlookErrorMapper.IsObjectModelGuardDenied(exception))
                        {
                            throw new OutlookGatewayException(OutlookGatewayFailure.ObjectModelGuard);
                        }
                        finally
                        {
                            if (recipient != null)
                            {
                                Marshal.ReleaseComObject(recipient);
                                OutlookComMetrics.RecordReleased();
                            }
                        }
                    }
                }

                return new OutlookMessageDetail(
                    summary,
                    toRecipients,
                    ccRecipients,
                    bccRecipients,
                    totalTo,
                    totalCc,
                    totalBcc,
                    ReadBody(mail, request));
            }
            catch (COMException exception) when (exception.ErrorCode == MapiNotSupported)
            {
                throw new OutlookGatewayException(OutlookGatewayFailure.UnsupportedStore);
            }
            finally
            {
                if (recipients != null)
                {
                    Marshal.ReleaseComObject(recipients);
                    OutlookComMetrics.RecordReleased();
                }

                if (rawItem != null && Marshal.IsComObject(rawItem))
                {
                    Marshal.ReleaseComObject(rawItem);
                    OutlookComMetrics.RecordReleased();
                }

                if (store != null)
                {
                    Marshal.ReleaseComObject(store);
                    OutlookComMetrics.RecordReleased();
                }
            }
        }

        public static OutlookMessagePage GetConversation(
            OutlookRuntimeContext context,
            OutlookGetConversationRequest request,
            CancellationToken cancellationToken)
        {
            context.AssertAndCaptureCurrentThread();
            cancellationToken.ThrowIfCancellationRequested();
            var session = RequireSession(context);

            Outlook.Store? store = null;
            object? rawSeed = null;
            Outlook.Conversation? conversation = null;
            Outlook.SimpleItems? roots = null;
            try
            {
                store = GetStore(session, request.Item.StoreId);
                if (!store.IsConversationEnabled)
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.UnsupportedStore);
                }

                rawSeed = GetItem(session, request.Item);
                if (!(rawSeed is Outlook.MailItem seed))
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.UnsupportedItemType);
                }

                RequireMatchingItemClass(seed, request.Item, OutlookGatewayFailure.ItemMovedOrDeleted);
                var seedFolder = OutlookReadProjection.ReadParentFolder(seed);
                var seedSummary = OutlookReadProjection.ProjectMessage(
                    seed,
                    seedFolder,
                    OutlookMessageTimestampKind.Automatic);
                var conversationId = seedSummary.ConversationId;

                if (request.Anchor != null)
                {
                    if (conversationId == null)
                    {
                        throw new OutlookGatewayException(OutlookGatewayFailure.CursorStale);
                    }

                    ValidateMessageAnchor(
                        session,
                        request.Anchor,
                        expectedFolder: null,
                        OutlookMessageTimestampKind.Automatic,
                        filter: null,
                        expectedConversationId: conversationId);
                }

                conversation = seed.GetConversation();
                if (conversation == null)
                {
                    if (request.Anchor != null)
                    {
                        throw new OutlookGatewayException(OutlookGatewayFailure.CursorStale);
                    }

                    var singleton = new List<OutlookMessageSummary> { seedSummary };
                    OutlookComMetrics.ObserveMaterializedItems(singleton.Count);
                    return OutlookReadProjection.BuildMessagePage(
                        singleton,
                        request.PageSize,
                        1,
                        Array.Empty<OutlookScopeFailure>());
                }
                OutlookComMetrics.RecordAcquired();

                roots = conversation.GetRootItems();
                if (roots != null)
                {
                    OutlookComMetrics.RecordAcquired();
                }
                var candidates = new List<OutlookMessageSummary>(request.PageSize + 1);
                if (roots != null)
                {
                    var examinedNodeCount = 0;
                    VisitConversationNodes(
                        conversation,
                        roots,
                        depth: 0,
                        request.Anchor,
                        request.PageSize,
                        candidates,
                        ref examinedNodeCount,
                        cancellationToken);
                }

                return OutlookReadProjection.BuildMessagePage(
                    candidates,
                    request.PageSize,
                    1,
                    Array.Empty<OutlookScopeFailure>());
            }
            catch (COMException exception) when (exception.ErrorCode == MapiNotSupported)
            {
                throw new OutlookGatewayException(OutlookGatewayFailure.UnsupportedStore);
            }
            finally
            {
                if (roots != null)
                {
                    Marshal.ReleaseComObject(roots);
                    OutlookComMetrics.RecordReleased();
                }

                if (conversation != null)
                {
                    Marshal.ReleaseComObject(conversation);
                    OutlookComMetrics.RecordReleased();
                }

                if (rawSeed != null && Marshal.IsComObject(rawSeed))
                {
                    Marshal.ReleaseComObject(rawSeed);
                    OutlookComMetrics.RecordReleased();
                }

                if (store != null)
                {
                    Marshal.ReleaseComObject(store);
                    OutlookComMetrics.RecordReleased();
                }
            }
        }

        public static OutlookAttachmentPage ListAttachments(
            OutlookRuntimeContext context,
            OutlookListAttachmentsRequest request,
            CancellationToken cancellationToken)
        {
            context.AssertAndCaptureCurrentThread();
            cancellationToken.ThrowIfCancellationRequested();
            var session = RequireSession(context);

            Outlook.Store? store = null;
            object? rawItem = null;
            Outlook.Attachments? attachments = null;
            try
            {
                store = GetStore(session, request.Item.StoreId);
                rawItem = GetItem(session, request.Item);
                if (!(rawItem is Outlook.MailItem mail))
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.UnsupportedItemType);
                }

                RequireMatchingItemClass(mail, request.Item, OutlookGatewayFailure.ItemMovedOrDeleted);
                attachments = mail.Attachments;
                if (attachments == null)
                {
                    return new OutlookAttachmentPage(
                        Array.Empty<OutlookAttachmentSummary>(),
                        nextAnchor: null);
                }
                OutlookComMetrics.RecordAcquired();

                if (request.Anchor != null)
                {
                    if (request.Anchor.AttachmentIndex > attachments.Count)
                    {
                        throw new OutlookGatewayException(OutlookGatewayFailure.CursorStale);
                    }

                    Outlook.Attachment? anchorAttachment = null;
                    try
                    {
                        anchorAttachment = attachments[request.Anchor.AttachmentIndex];
                        if (anchorAttachment == null)
                        {
                            throw new OutlookGatewayException(OutlookGatewayFailure.CursorStale);
                        }
                        OutlookComMetrics.RecordAcquired();
                        var projectedAnchor = ProjectAttachment(
                            anchorAttachment,
                            request.Item,
                            request.Anchor.AttachmentIndex);
                        if (!string.Equals(
                            projectedAnchor.Attachment.MetadataFingerprint,
                            request.Anchor.MetadataFingerprint,
                            StringComparison.Ordinal))
                        {
                            throw new OutlookGatewayException(OutlookGatewayFailure.CursorStale);
                        }
                    }
                    finally
                    {
                        if (anchorAttachment != null)
                        {
                            Marshal.ReleaseComObject(anchorAttachment);
                            OutlookComMetrics.RecordReleased();
                        }
                    }
                }

                var firstIndex = request.Anchor?.AttachmentIndex + 1 ?? 1;
                var maximumIndex = Math.Min(
                    attachments.Count,
                    firstIndex + request.PageSize);
                var processedExtraSourceRow =
                    maximumIndex == firstIndex + request.PageSize;
                var candidates = new List<OutlookAttachmentSummary>(request.PageSize + 1);
                for (var index = firstIndex; index <= maximumIndex; index++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    Outlook.Attachment? attachment = null;
                    try
                    {
                        attachment = attachments[index];
                        if (attachment == null)
                        {
                            throw new OutlookGatewayException(OutlookGatewayFailure.Internal);
                        }
                        OutlookComMetrics.RecordAcquired();

                        candidates.Add(ProjectAttachment(attachment, request.Item, index));
                        OutlookComMetrics.ObserveMaterializedItems(candidates.Count);
                    }
                    finally
                    {
                        if (attachment != null)
                        {
                            Marshal.ReleaseComObject(attachment);
                            OutlookComMetrics.RecordReleased();
                        }
                    }
                }

                var hasMore = processedExtraSourceRow;
                if (hasMore)
                {
                    candidates.RemoveAt(candidates.Count - 1);
                }

                OutlookAttachmentKeysetAnchor? nextAnchor = null;
                if (hasMore && candidates.Count > 0)
                {
                    var last = candidates[candidates.Count - 1].Attachment;
                    nextAnchor = new OutlookAttachmentKeysetAnchor(
                        last.AttachmentIndex,
                        last.MetadataFingerprint);
                }

                return new OutlookAttachmentPage(candidates, nextAnchor);
            }
            catch (COMException exception) when (exception.ErrorCode == MapiNotSupported)
            {
                throw new OutlookGatewayException(OutlookGatewayFailure.UnsupportedStore);
            }
            finally
            {
                if (attachments != null)
                {
                    Marshal.ReleaseComObject(attachments);
                    OutlookComMetrics.RecordReleased();
                }

                if (rawItem != null && Marshal.IsComObject(rawItem))
                {
                    Marshal.ReleaseComObject(rawItem);
                    OutlookComMetrics.RecordReleased();
                }

                if (store != null)
                {
                    Marshal.ReleaseComObject(store);
                    OutlookComMetrics.RecordReleased();
                }
            }
        }

        private static Outlook.NameSpace RequireSession(OutlookRuntimeContext context)
        {
            var application = context.Application;
            var session = application.Session;
            return session ?? throw new OutlookGatewayException(OutlookGatewayFailure.NotReady);
        }

        private static Outlook.Store GetStore(Outlook.NameSpace session, string storeId)
        {
            try
            {
                var store = session.GetStoreFromID(storeId) ??
                    throw new OutlookGatewayException(OutlookGatewayFailure.StoreNotFound);
                OutlookComMetrics.RecordAcquired();
                return store;
            }
            catch (Exception exception)
            {
                throw OutlookErrorMapper.MapLookup(
                    exception,
                    OutlookGatewayFailure.StoreNotFound);
            }
        }

        private static void RejectVisibleSearchFolder(
            Outlook.NameSpace session,
            Outlook.Store store,
            Outlook.MAPIFolder folder,
            CancellationToken cancellationToken)
        {
            Outlook.Folders? searchFolders = null;
            try
            {
                searchFolders = store.GetSearchFolders();
                if (searchFolders == null)
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.UnsupportedStore);
                }
                OutlookComMetrics.RecordAcquired();

                var searchFolderCount = searchFolders.Count;
                OutlookSearchFolderPolicy.RequireBoundedCount(searchFolderCount);
                var targetEntryId = OutlookReadProjection.RequireIdentifier(
                    folder.EntryID,
                    FolderRef.MaximumEntryIdLength);
                for (var index = 1; index <= searchFolderCount; index++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    Outlook.MAPIFolder? searchFolder = null;
                    try
                    {
                        searchFolder = searchFolders[index];
                        if (searchFolder == null)
                        {
                            throw new OutlookGatewayException(
                                OutlookGatewayFailure.UnsupportedStore);
                        }
                        OutlookComMetrics.RecordAcquired();

                        var searchEntryId = OutlookReadProjection.RequireIdentifier(
                            searchFolder.EntryID,
                            FolderRef.MaximumEntryIdLength);
                        OutlookSearchFolderPolicy.RejectIfMatch(
                            session.CompareEntryIDs(searchEntryId, targetEntryId));
                    }
                    finally
                    {
                        if (searchFolder != null)
                        {
                            Marshal.ReleaseComObject(searchFolder);
                            OutlookComMetrics.RecordReleased();
                        }
                    }
                }
            }
            catch (COMException exception)
                when (exception.ErrorCode == unchecked((int)0x8004010F) ||
                    exception.ErrorCode == MapiNotSupported)
            {
                throw new OutlookGatewayException(OutlookGatewayFailure.UnsupportedStore);
            }
            finally
            {
                if (searchFolders != null)
                {
                    Marshal.ReleaseComObject(searchFolders);
                    OutlookComMetrics.RecordReleased();
                }
            }
        }

        private static Outlook.MAPIFolder GetFolder(
            Outlook.NameSpace session,
            FolderRef folder)
        {
            try
            {
                var result = session.GetFolderFromID(folder.EntryId, folder.StoreId) ??
                    throw new OutlookGatewayException(OutlookGatewayFailure.FolderNotFound);
                OutlookComMetrics.RecordAcquired();
                return result;
            }
            catch (Exception exception)
            {
                throw OutlookErrorMapper.MapLookup(
                    exception,
                    OutlookGatewayFailure.FolderNotFound);
            }
        }

        private static Outlook.MAPIFolder GetRootFolder(Outlook.Store store)
        {
            try
            {
                var folder = store.GetRootFolder() ??
                    throw new OutlookGatewayException(OutlookGatewayFailure.UnsupportedStore);
                OutlookComMetrics.RecordAcquired();
                return folder;
            }
            catch (Exception exception)
            {
                throw OutlookErrorMapper.MapLookup(
                    exception,
                    OutlookGatewayFailure.UnsupportedStore);
            }
        }

        private static Outlook.MAPIFolder GetDefaultFolder(
            Outlook.Store store,
            Outlook.OlDefaultFolders folderKind)
        {
            try
            {
                var folder = store.GetDefaultFolder(folderKind) ??
                    throw new OutlookGatewayException(OutlookGatewayFailure.FolderNotFound);
                OutlookComMetrics.RecordAcquired();
                return folder;
            }
            catch (Exception exception)
            {
                throw OutlookErrorMapper.MapLookup(
                    exception,
                    OutlookGatewayFailure.FolderNotFound);
            }
        }

        private static object GetItem(Outlook.NameSpace session, ItemRef item)
        {
            try
            {
                var result = session.GetItemFromID(item.EntryId, item.StoreId) ??
                    throw new OutlookGatewayException(OutlookGatewayFailure.ItemNotFound);
                if (Marshal.IsComObject(result))
                {
                    OutlookComMetrics.RecordAcquired();
                }

                return result;
            }
            catch (Exception exception)
            {
                throw OutlookErrorMapper.MapLookup(
                    exception,
                    OutlookGatewayFailure.ItemNotFound);
            }
        }

        private static OutlookMailboxSummary ProjectMailbox(Outlook.Store store)
        {
            var storeId = OutlookReadProjection.RequireIdentifier(
                store.StoreID,
                MailboxRef.MaximumStoreIdLength);
            var storeType = ReadStoreType(store);
            var capabilities = new OutlookStoreCapabilities(
                storeType != OutlookStoreType.NonExchange &&
                    storeType != OutlookStoreType.Unknown,
                ReadStoreBoolean(store, static value => value.IsDataFileStore),
                ReadStoreBoolean(store, static value => value.IsCachedExchange));
            return new OutlookMailboxSummary(
                new MailboxRef(storeId),
                OutlookReadProjection.BoundDisplay(
                    store.DisplayName,
                    OutlookMailboxSummary.MaximumDisplayNameLength),
                storeType,
                capabilities,
                new OutlookStandardFolderReferences(
                    TryGetDefaultFolderRef(store, Outlook.OlDefaultFolders.olFolderInbox),
                    TryGetDefaultFolderRef(store, Outlook.OlDefaultFolders.olFolderDrafts),
                    TryGetDefaultFolderRef(store, Outlook.OlDefaultFolders.olFolderSentMail),
                    TryGetDefaultFolderRef(store, Outlook.OlDefaultFolders.olFolderDeletedItems),
                    archive: null));
        }

        private static OutlookStoreType ReadStoreType(Outlook.Store store)
        {
            try
            {
                return OutlookReadProjection.MapStoreType(store.ExchangeStoreType);
            }
            catch (COMException exception) when (exception.ErrorCode == MapiNotSupported)
            {
                return OutlookStoreType.Unknown;
            }
        }

        private static bool ReadStoreBoolean(
            Outlook.Store store,
            Func<Outlook.Store, bool> reader)
        {
            try
            {
                return reader(store);
            }
            catch (COMException exception) when (exception.ErrorCode == MapiNotSupported)
            {
                return false;
            }
        }

        private static FolderRef? TryGetDefaultFolderRef(
            Outlook.Store store,
            Outlook.OlDefaultFolders folderKind)
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
                    ? null
                    : new FolderRef(
                        OutlookReadProjection.RequireIdentifier(
                            folder.StoreID,
                            MailboxRef.MaximumStoreIdLength),
                        OutlookReadProjection.RequireIdentifier(
                            folder.EntryID,
                            FolderRef.MaximumEntryIdLength));
            }
            catch (COMException exception)
                when (exception.ErrorCode == unchecked((int)0x8004010F) ||
                    exception.ErrorCode == unchecked((int)0x80040102))
            {
                return null;
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

        private static void ValidateMailboxAnchor(
            Outlook.NameSpace session,
            OutlookMailboxKeysetAnchor anchor)
        {
            Outlook.Store? store = null;
            try
            {
                store = GetStore(session, anchor.Mailbox.StoreId);
                if (!string.Equals(
                    OutlookReadProjection.BoundDisplay(
                        store.DisplayName,
                        OutlookMailboxSummary.MaximumDisplayNameLength),
                    anchor.DisplayName,
                    StringComparison.Ordinal))
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.CursorStale);
                }
            }
            catch (OutlookGatewayException exception)
                when (exception.Failure == OutlookGatewayFailure.StoreNotFound)
            {
                throw new OutlookGatewayException(OutlookGatewayFailure.CursorStale);
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

        private static void ValidateFolderAnchor(
            Outlook.NameSpace session,
            Outlook.MAPIFolder expectedParent,
            OutlookFolderKeysetAnchor anchor)
        {
            Outlook.MAPIFolder? folder = null;
            object? parent = null;
            try
            {
                folder = GetFolder(session, anchor.Folder);
                if (!string.Equals(
                    OutlookReadProjection.BoundDisplay(
                        folder.Name,
                        OutlookFolderSummary.MaximumDisplayNameLength),
                    anchor.DisplayName,
                    StringComparison.Ordinal))
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.CursorStale);
                }

                parent = folder.Parent;
                if (parent != null && Marshal.IsComObject(parent))
                {
                    OutlookComMetrics.RecordAcquired();
                }

                if (!(parent is Outlook.MAPIFolder parentFolder) ||
                    !string.Equals(parentFolder.StoreID, expectedParent.StoreID, StringComparison.Ordinal) ||
                    !string.Equals(parentFolder.EntryID, expectedParent.EntryID, StringComparison.Ordinal))
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.CursorStale);
                }
            }
            catch (OutlookGatewayException exception)
                when (exception.Failure == OutlookGatewayFailure.FolderNotFound)
            {
                throw new OutlookGatewayException(OutlookGatewayFailure.CursorStale);
            }
            finally
            {
                if (parent != null && Marshal.IsComObject(parent))
                {
                    Marshal.ReleaseComObject(parent);
                    OutlookComMetrics.RecordReleased();
                }

                if (folder != null)
                {
                    Marshal.ReleaseComObject(folder);
                    OutlookComMetrics.RecordReleased();
                }
            }
        }

        private static OutlookMessageTimestampKind ResolveTimestampKind(
            Outlook.NameSpace session,
            Outlook.Store store,
            Outlook.MAPIFolder folder)
        {
            Outlook.MAPIFolder? sent = null;
            try
            {
                sent = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);
                if (sent != null)
                {
                    OutlookComMetrics.RecordAcquired();
                }

                if (sent != null && session.CompareEntryIDs(sent.EntryID, folder.EntryID))
                {
                    return OutlookMessageTimestampKind.Sent;
                }
            }
            catch (COMException exception)
                when (exception.ErrorCode == unchecked((int)0x8004010F) ||
                    exception.ErrorCode == unchecked((int)0x80040102))
            {
            }
            finally
            {
                if (sent != null)
                {
                    Marshal.ReleaseComObject(sent);
                    OutlookComMetrics.RecordReleased();
                }
            }

            Outlook.MAPIFolder? drafts = null;
            try
            {
                drafts = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts);
                if (drafts != null)
                {
                    OutlookComMetrics.RecordAcquired();
                }

                if (drafts != null && session.CompareEntryIDs(drafts.EntryID, folder.EntryID))
                {
                    return OutlookMessageTimestampKind.Modified;
                }
            }
            catch (COMException exception)
                when (exception.ErrorCode == unchecked((int)0x8004010F) ||
                    exception.ErrorCode == unchecked((int)0x80040102))
            {
            }
            finally
            {
                if (drafts != null)
                {
                    Marshal.ReleaseComObject(drafts);
                    OutlookComMetrics.RecordReleased();
                }
            }

            return OutlookMessageTimestampKind.Received;
        }

        private static List<OutlookMessageSummary> ReadMessagesFromFolder(
            Outlook.MAPIFolder folder,
            FolderRef folderRef,
            string restriction,
            OutlookMessageTimestampKind timestampKind,
            OutlookMessageKeysetAnchor? anchor,
            OutlookMessageSearchFilter? filter,
            int pageSize,
            CancellationToken cancellationToken)
        {
            Outlook.Items? allItems = null;
            Outlook.Items? restrictedItems = null;
            try
            {
                allItems = folder.Items;
                if (allItems == null)
                {
                    return new List<OutlookMessageSummary>();
                }
                OutlookComMetrics.RecordAcquired();

                restrictedItems = allItems.Restrict(restriction);
                if (restrictedItems == null)
                {
                    return new List<OutlookMessageSummary>();
                }
                OutlookComMetrics.RecordAcquired();

                restrictedItems.Sort(OutlookReadFilter.SortProperty(timestampKind), true);
                var candidates = new List<OutlookMessageSummary>(pageSize + 1);
                var examinedCount = 0;
                var timestampTieCount = 0;
                DateTime? priorTimestamp = null;
                object? current = null;
                try
                {
                    current = restrictedItems.GetFirst();
                    if (current != null && Marshal.IsComObject(current))
                    {
                        OutlookComMetrics.RecordAcquired();
                    }

                    while (current != null)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        examinedCount++;
                        if (examinedCount > MaximumMessagesExamined)
                        {
                            throw new OutlookGatewayException(OutlookGatewayFailure.Timeout);
                        }

                        var currentTimestamp = DateTime.SpecifyKind(DateTime.MinValue, DateTimeKind.Utc);
                        if (current is Outlook.MailItem mail)
                        {
                            currentTimestamp = OutlookReadProjection.ReadOrderingTimestamp(
                                mail,
                                timestampKind);
                            if (priorTimestamp.HasValue &&
                                priorTimestamp.Value == currentTimestamp)
                            {
                                timestampTieCount++;
                            }
                            else
                            {
                                priorTimestamp = currentTimestamp;
                                timestampTieCount = 1;
                            }

                            if (timestampTieCount > MaximumMessageTimestampTieGroup)
                            {
                                throw new OutlookGatewayException(OutlookGatewayFailure.Timeout);
                            }

                            if (filter == null || OutlookReadFilter.Matches(mail, filter))
                            {
                                var summary = OutlookReadProjection.ProjectMessage(
                                    mail,
                                    folderRef,
                                    timestampKind);
                                if (anchor == null ||
                                    OutlookReadProjection.CompareMessageToAnchor(summary, anchor) > 0)
                                {
                                    OutlookReadProjection.InsertBounded(
                                        candidates,
                                        summary,
                                        pageSize + 1,
                                        OutlookReadProjection.CompareMessages);
                                    OutlookComMetrics.ObserveMaterializedItems(candidates.Count);
                                }
                            }
                        }

                        if (Marshal.IsComObject(current))
                        {
                            Marshal.ReleaseComObject(current);
                            OutlookComMetrics.RecordReleased();
                        }
                        current = null;
                        if (candidates.Count == pageSize + 1 &&
                            currentTimestamp < candidates[candidates.Count - 1].EffectiveTimestampUtc)
                        {
                            break;
                        }

                        current = restrictedItems.GetNext();
                        if (current != null && Marshal.IsComObject(current))
                        {
                            OutlookComMetrics.RecordAcquired();
                        }
                    }
                }
                finally
                {
                    if (current != null && Marshal.IsComObject(current))
                    {
                        Marshal.ReleaseComObject(current);
                        OutlookComMetrics.RecordReleased();
                    }
                }

                return candidates;
            }
            catch (COMException exception) when (exception.ErrorCode == MapiNotSupported)
            {
                throw new OutlookGatewayException(OutlookGatewayFailure.UnsupportedStore);
            }
            finally
            {
                if (restrictedItems != null)
                {
                    Marshal.ReleaseComObject(restrictedItems);
                    OutlookComMetrics.RecordReleased();
                }

                if (allItems != null)
                {
                    Marshal.ReleaseComObject(allItems);
                    OutlookComMetrics.RecordReleased();
                }
            }
        }

        private static void ValidateMessageAnchor(
            Outlook.NameSpace session,
            OutlookMessageKeysetAnchor anchor,
            FolderRef? expectedFolder,
            OutlookMessageTimestampKind timestampKind,
            OutlookMessageSearchFilter? filter,
            string? expectedConversationId)
        {
            object? rawItem = null;
            try
            {
                rawItem = GetItem(session, anchor.Item);
                if (!(rawItem is Outlook.MailItem mail))
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.CursorStale);
                }

                RequireMatchingItemClass(mail, anchor.Item, OutlookGatewayFailure.CursorStale);
                var parent = OutlookReadProjection.ReadParentFolder(mail);
                if (expectedFolder != null && !SameFolder(parent, expectedFolder))
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.CursorStale);
                }

                if (filter != null && !OutlookReadFilter.Matches(mail, filter))
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.CursorStale);
                }

                var projected = OutlookReadProjection.ProjectMessage(mail, parent, timestampKind);
                if (projected.EffectiveTimestampUtc != anchor.EffectiveTimestampUtc ||
                    (expectedConversationId != null &&
                    !string.Equals(
                        projected.ConversationId,
                        expectedConversationId,
                        StringComparison.Ordinal)))
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.CursorStale);
                }
            }
            catch (OutlookGatewayException exception)
                when (exception.Failure == OutlookGatewayFailure.ItemNotFound)
            {
                throw new OutlookGatewayException(OutlookGatewayFailure.CursorStale);
            }
            finally
            {
                if (rawItem != null && Marshal.IsComObject(rawItem))
                {
                    Marshal.ReleaseComObject(rawItem);
                    OutlookComMetrics.RecordReleased();
                }
            }
        }

        private static void RequireMatchingItemClass(
            Outlook.MailItem mail,
            ItemRef expected,
            OutlookGatewayFailure failure)
        {
            if (!string.Equals(mail.MessageClass, expected.ItemClass, StringComparison.Ordinal))
            {
                throw new OutlookGatewayException(failure);
            }
        }

        private static OutlookMessageBody ReadBody(
            Outlook.MailItem mail,
            OutlookGetMessageRequest request)
        {
            string content;
            try
            {
                content = request.BodyFormat == OutlookBodyFormat.Html
                    ? mail.HTMLBody
                    : mail.Body;
            }
            catch (COMException exception)
                when (OutlookErrorMapper.IsAccessDenied(exception))
            {
                return new OutlookMessageBody(
                    request.BodyFormat,
                    string.Empty,
                    originalCharacterCount: null,
                    isTruncated: false,
                    isProtected: true);
            }
            catch (COMException exception)
                when (OutlookErrorMapper.IsObjectModelGuardDenied(exception))
            {
                throw new OutlookGatewayException(OutlookGatewayFailure.ObjectModelGuard);
            }

            content = content ?? string.Empty;
            var originalLength = content.Length;
            if (content.Length > request.MaximumBodyCharacters)
            {
                content = content.Substring(0, request.MaximumBodyCharacters);
            }

            return new OutlookMessageBody(
                request.BodyFormat,
                content,
                originalLength,
                content.Length < originalLength,
                isProtected: false);
        }

        private static void VisitConversationNodes(
            Outlook.Conversation conversation,
            Outlook.SimpleItems nodes,
            int depth,
            OutlookMessageKeysetAnchor? anchor,
            int pageSize,
            List<OutlookMessageSummary> candidates,
            ref int examinedNodeCount,
            CancellationToken cancellationToken)
        {
            if (depth > MaximumConversationDepth)
            {
                throw new OutlookGatewayException(OutlookGatewayFailure.Internal);
            }

            for (var index = 1; index <= nodes.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                examinedNodeCount++;
                if (examinedNodeCount > MaximumConversationNodesExamined)
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.Timeout);
                }

                object? node = null;
                Outlook.SimpleItems? children = null;
                try
                {
                    node = nodes[index];
                    if (node == null)
                    {
                        continue;
                    }
                    if (Marshal.IsComObject(node))
                    {
                        OutlookComMetrics.RecordAcquired();
                    }

                    if (node is Outlook.MailItem mail)
                    {
                        var folder = OutlookReadProjection.ReadParentFolder(mail);
                        var summary = OutlookReadProjection.ProjectMessage(
                            mail,
                            folder,
                            OutlookMessageTimestampKind.Automatic);
                        if (anchor == null ||
                            OutlookReadProjection.CompareMessageToAnchor(summary, anchor) > 0)
                        {
                            OutlookReadProjection.InsertBounded(
                                candidates,
                                summary,
                                pageSize + 1,
                                OutlookReadProjection.CompareMessages);
                            OutlookComMetrics.ObserveMaterializedItems(candidates.Count);
                        }
                    }

                    children = conversation.GetChildren(node);
                    if (children != null)
                    {
                        OutlookComMetrics.RecordAcquired();
                    }

                    if (children != null && children.Count > 0)
                    {
                        VisitConversationNodes(
                            conversation,
                            children,
                            depth + 1,
                            anchor,
                            pageSize,
                            candidates,
                            ref examinedNodeCount,
                            cancellationToken);
                    }
                }
                finally
                {
                    if (children != null)
                    {
                        Marshal.ReleaseComObject(children);
                        OutlookComMetrics.RecordReleased();
                    }

                    if (node != null && Marshal.IsComObject(node))
                    {
                        Marshal.ReleaseComObject(node);
                        OutlookComMetrics.RecordReleased();
                    }
                }
            }
        }

        private static OutlookAttachmentSummary ProjectAttachment(
            Outlook.Attachment attachment,
            ItemRef item,
            int index)
        {
            Outlook.PropertyAccessor? accessor = null;
            var fileName = OutlookReadProjection.BoundOptionalText(
                attachment.FileName,
                AttachmentRef.MaximumNameLength);
            var displayName = OutlookReadProjection.BoundOptionalText(
                attachment.DisplayName,
                AttachmentRef.MaximumNameLength);
            var name = fileName ?? displayName ?? "attachment-" + index;
            var size = (long)attachment.Size;
            var type = (int)attachment.Type;
            var position = attachment.Position;
            var blockLevel = (int)attachment.BlockLevel;
            string? contentType = null;
            try
            {
                accessor = attachment.PropertyAccessor;
                if (accessor != null)
                {
                    OutlookComMetrics.RecordAcquired();
                    contentType = OutlookReadProjection.BoundOptionalText(
                        accessor.GetProperty(AttachmentMimeTagSchema) as string,
                        OutlookAttachmentSummary.MaximumContentTypeLength);
                }
            }
            catch (COMException exception)
                when (exception.ErrorCode == unchecked((int)0x8004010F) ||
                    exception.ErrorCode == unchecked((int)0x80040102) ||
                    exception.ErrorCode == unchecked((int)0x80070005))
            {
                contentType = null;
            }
            finally
            {
                if (accessor != null)
                {
                    Marshal.ReleaseComObject(accessor);
                    OutlookComMetrics.RecordReleased();
                }
            }

            var fingerprint = OutlookReadProjection.ComputeAttachmentFingerprint(
                item,
                index,
                name,
                displayName,
                size,
                type,
                position,
                blockLevel,
                contentType);
            return new OutlookAttachmentSummary(
                new AttachmentRef(
                    item,
                    index,
                    name,
                    size,
                    sizeIsKnown: size != 0,
                    metadataFingerprint: fingerprint),
                contentType);
        }

        private static OutlookMailboxPage BuildMailboxPage(
            List<OutlookMailboxSummary> candidates,
            int pageSize)
        {
            var hasMore = candidates.Count > pageSize;
            if (hasMore)
            {
                candidates.RemoveAt(candidates.Count - 1);
            }

            OutlookMailboxKeysetAnchor? nextAnchor = null;
            if (hasMore && candidates.Count > 0)
            {
                var last = candidates[candidates.Count - 1];
                nextAnchor = new OutlookMailboxKeysetAnchor(last.DisplayName, last.Mailbox);
            }

            return new OutlookMailboxPage(candidates, nextAnchor);
        }

        private static OutlookFolderPage BuildFolderPage(
            List<OutlookFolderSummary> candidates,
            int pageSize)
        {
            var hasMore = candidates.Count > pageSize;
            if (hasMore)
            {
                candidates.RemoveAt(candidates.Count - 1);
            }

            OutlookFolderKeysetAnchor? nextAnchor = null;
            if (hasMore && candidates.Count > 0)
            {
                var last = candidates[candidates.Count - 1];
                nextAnchor = new OutlookFolderKeysetAnchor(last.DisplayName, last.Folder);
            }

            return new OutlookFolderPage(candidates, nextAnchor);
        }

        private static bool SameFolder(FolderRef left, FolderRef right)
        {
            return string.Equals(left.StoreId, right.StoreId, StringComparison.Ordinal) &&
                string.Equals(left.EntryId, right.EntryId, StringComparison.Ordinal);
        }
    }
}
