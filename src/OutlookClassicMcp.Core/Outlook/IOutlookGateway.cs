using System.Threading;
using System.Threading.Tasks;

namespace OutlookClassicMcp.Core.Outlook
{
    public interface IOutlookGateway
    {
        Task<OutlookProbeSnapshot> ProbeAsync(CancellationToken cancellationToken);

        Task<OutlookMailboxPage> ListMailboxesAsync(
            OutlookListMailboxesRequest request,
            CancellationToken cancellationToken);

        Task<OutlookFolderPage> ListFoldersAsync(
            OutlookListFoldersRequest request,
            CancellationToken cancellationToken);

        Task<OutlookMessagePage> ListMessagesAsync(
            OutlookListMessagesRequest request,
            CancellationToken cancellationToken);

        Task<OutlookMessagePage> SearchMessagesAsync(
            OutlookSearchMessagesRequest request,
            CancellationToken cancellationToken);

        Task<OutlookMessageDetail> GetMessageAsync(
            OutlookGetMessageRequest request,
            CancellationToken cancellationToken);

        Task<OutlookMessagePage> GetConversationAsync(
            OutlookGetConversationRequest request,
            CancellationToken cancellationToken);

        Task<OutlookAttachmentPage> ListAttachmentsAsync(
            OutlookListAttachmentsRequest request,
            CancellationToken cancellationToken);
    }
}
