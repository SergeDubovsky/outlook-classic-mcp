using System;

namespace OutlookClassicMcp.Core.Outlook
{
    public enum OutlookGatewayFailure
    {
        NotReady = 0,
        Degraded = 1,
        Stopping = 2,
        QueueFull = 3,
        Timeout = 4,
        ComBusy = 5,
        AccessDenied = 6,
        ObjectModelGuard = 7,
        StaDispatchFailed = 8,
        Internal = 9,
        Paused = 10,
        StoreNotFound = 11,
        FolderNotFound = 12,
        ItemNotFound = 13,
        ItemMovedOrDeleted = 14,
        UnsupportedStore = 15,
        UnsupportedItemType = 16,
        InvalidArgument = 17,
        CursorStale = 18,
    }

    public sealed class OutlookGatewayException : Exception
    {
        public OutlookGatewayException(OutlookGatewayFailure failure)
            : base(GetSafeMessage(failure))
        {
            Failure = failure;
        }

        public OutlookGatewayFailure Failure { get; }

        private static string GetSafeMessage(OutlookGatewayFailure failure)
        {
            switch (failure)
            {
                case OutlookGatewayFailure.NotReady:
                    return "Outlook is not ready.";
                case OutlookGatewayFailure.Degraded:
                    return "The Outlook integration is degraded.";
                case OutlookGatewayFailure.Stopping:
                    return "The Outlook integration is stopping.";
                case OutlookGatewayFailure.QueueFull:
                    return "The Outlook request queue is full.";
                case OutlookGatewayFailure.Timeout:
                    return "The Outlook request timed out.";
                case OutlookGatewayFailure.ComBusy:
                    return "Outlook is busy.";
                case OutlookGatewayFailure.AccessDenied:
                    return "Outlook denied access.";
                case OutlookGatewayFailure.ObjectModelGuard:
                    return "Outlook blocked the operation.";
                case OutlookGatewayFailure.StaDispatchFailed:
                    return "The Outlook UI thread dispatch failed.";
                case OutlookGatewayFailure.Internal:
                    return "The Outlook operation failed.";
                case OutlookGatewayFailure.Paused:
                    return "The Outlook integration is paused.";
                case OutlookGatewayFailure.StoreNotFound:
                    return "The mailbox store no longer exists.";
                case OutlookGatewayFailure.FolderNotFound:
                    return "The folder no longer exists.";
                case OutlookGatewayFailure.ItemNotFound:
                    return "The message no longer exists.";
                case OutlookGatewayFailure.ItemMovedOrDeleted:
                    return "The message was moved or deleted.";
                case OutlookGatewayFailure.UnsupportedStore:
                    return "The mailbox store is not supported for this operation.";
                case OutlookGatewayFailure.UnsupportedItemType:
                    return "The Outlook item type is not supported.";
                case OutlookGatewayFailure.InvalidArgument:
                    return "The request is invalid.";
                case OutlookGatewayFailure.CursorStale:
                    return "The continuation cursor is stale.";
                default:
                    throw new ArgumentOutOfRangeException(nameof(failure));
            }
        }
    }
}
