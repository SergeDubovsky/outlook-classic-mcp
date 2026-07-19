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
                default:
                    throw new ArgumentOutOfRangeException(nameof(failure));
            }
        }
    }
}
