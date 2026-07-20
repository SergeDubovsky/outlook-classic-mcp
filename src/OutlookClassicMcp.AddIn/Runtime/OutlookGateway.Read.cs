using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OutlookClassicMcp.Core.Outlook;

namespace OutlookClassicMcp.AddIn.Runtime
{
    internal sealed partial class OutlookGateway
    {
        public Task<OutlookMailboxPage> ListMailboxesAsync(
            OutlookListMailboxesRequest request,
            CancellationToken cancellationToken)
        {
            return DispatchAsync(
                RequireRequest(request, nameof(request)),
                OutlookReadOperations.ListMailboxes,
                cancellationToken);
        }

        public Task<OutlookFolderPage> ListFoldersAsync(
            OutlookListFoldersRequest request,
            CancellationToken cancellationToken)
        {
            return DispatchAsync(
                RequireRequest(request, nameof(request)),
                OutlookReadOperations.ListFolders,
                cancellationToken);
        }

        public Task<OutlookMessagePage> ListMessagesAsync(
            OutlookListMessagesRequest request,
            CancellationToken cancellationToken)
        {
            return DispatchAsync(
                RequireRequest(request, nameof(request)),
                OutlookReadOperations.ListMessages,
                cancellationToken);
        }

        public async Task<OutlookMessagePage> SearchMessagesAsync(
            OutlookSearchMessagesRequest request,
            CancellationToken cancellationToken)
        {
            RequireRequest(request, nameof(request));
            try
            {
                var scopes = OutlookSearchMerge.CanonicalizeScopes(request.Scopes);
                if (request.Anchor != null)
                {
                    await _dispatcher.InvokeWithContextAsync<
                            OutlookRuntimeContext,
                            OutlookSearchAnchorState,
                            bool>(
                            new OutlookSearchAnchorState(request, scopes),
                            OutlookReadOperations.ValidateSearchAnchor,
                            cancellationToken)
                        .ConfigureAwait(false);
                }

                var merge = new OutlookSearchMerge(
                    request.PageSize,
                    request.Scopes.Count);

                foreach (var scope in scopes)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    try
                    {
                        var scopeItems = await _dispatcher.InvokeWithContextAsync<
                                OutlookRuntimeContext,
                                OutlookSearchScopeState,
                                List<OutlookMessageSummary>>(
                                new OutlookSearchScopeState(
                                    scope,
                                    request.Filter,
                                    request.Anchor,
                                    request.PageSize),
                                OutlookReadOperations.SearchScope,
                                cancellationToken)
                            .ConfigureAwait(false);
                        merge.AddSuccess(scopeItems);
                    }
                    catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                    {
                        throw;
                    }
                    catch (Exception exception)
                    {
                        merge.AddFailure(scope, exception);
                    }
                }

                return merge.Complete();
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

        public Task<OutlookMessageDetail> GetMessageAsync(
            OutlookGetMessageRequest request,
            CancellationToken cancellationToken)
        {
            return DispatchAsync(
                RequireRequest(request, nameof(request)),
                OutlookReadOperations.GetMessage,
                cancellationToken);
        }

        public Task<OutlookMessagePage> GetConversationAsync(
            OutlookGetConversationRequest request,
            CancellationToken cancellationToken)
        {
            return DispatchAsync(
                RequireRequest(request, nameof(request)),
                OutlookReadOperations.GetConversation,
                cancellationToken);
        }

        public Task<OutlookAttachmentPage> ListAttachmentsAsync(
            OutlookListAttachmentsRequest request,
            CancellationToken cancellationToken)
        {
            return DispatchAsync(
                RequireRequest(request, nameof(request)),
                OutlookReadOperations.ListAttachments,
                cancellationToken);
        }

        private async Task<TResult> DispatchAsync<TState, TResult>(
            TState state,
            Func<OutlookRuntimeContext, TState, CancellationToken, TResult> operation,
            CancellationToken cancellationToken)
        {
            try
            {
                return await _dispatcher.InvokeWithContextAsync<
                        OutlookRuntimeContext,
                        TState,
                        TResult>(state, operation, cancellationToken)
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

        private static T RequireRequest<T>(T? request, string parameterName)
            where T : class
        {
            return request ?? throw new ArgumentNullException(parameterName);
        }

    }
}
