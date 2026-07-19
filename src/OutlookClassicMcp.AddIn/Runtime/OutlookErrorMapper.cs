using System;
using System.Runtime.InteropServices;
using OutlookClassicMcp.Core.Outlook;

namespace OutlookClassicMcp.AddIn.Runtime
{
    internal static class OutlookErrorMapper
    {
        private const int RpcCallRejected = unchecked((int)0x80010001);
        private const int RpcDisconnected = unchecked((int)0x80010108);
        private const int RpcServerCallRetryLater = unchecked((int)0x8001010A);
        private const int AccessDenied = unchecked((int)0x80070005);
        private const int ObjectNotConnected = unchecked((int)0x800401FD);

        public static OutlookGatewayException Map(Exception exception)
        {
            if (exception == null)
            {
                throw new ArgumentNullException(nameof(exception));
            }

            if (exception is OutlookGatewayException mapped)
            {
                return mapped;
            }

            if (exception is InvalidOperationException invalidOperation)
            {
                switch (invalidOperation.Message)
                {
                    case "HOST_BUSY":
                        return new OutlookGatewayException(OutlookGatewayFailure.QueueFull);
                    case "HOST_STOPPING":
                        return new OutlookGatewayException(OutlookGatewayFailure.Stopping);
                    case "HOST_UNAVAILABLE":
                    case "Outlook work executed outside the captured UI STA.":
                        return new OutlookGatewayException(OutlookGatewayFailure.StaDispatchFailed);
                }
            }

            if (exception is COMException comException)
            {
                switch (comException.ErrorCode)
                {
                    case RpcCallRejected:
                    case RpcServerCallRetryLater:
                        return new OutlookGatewayException(OutlookGatewayFailure.ComBusy);
                    case AccessDenied:
                        return new OutlookGatewayException(OutlookGatewayFailure.AccessDenied);
                    case RpcDisconnected:
                    case ObjectNotConnected:
                        return new OutlookGatewayException(OutlookGatewayFailure.NotReady);
                }
            }

            return new OutlookGatewayException(OutlookGatewayFailure.Internal);
        }
    }
}
