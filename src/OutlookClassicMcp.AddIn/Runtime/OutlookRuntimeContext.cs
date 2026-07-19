using System;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookClassicMcp.AddIn.Runtime
{
    internal sealed class OutlookRuntimeContext
    {
        public OutlookRuntimeContext(
            Outlook.Application application,
            OutlookThreadContext capturedThread)
        {
            Application = application ?? throw new ArgumentNullException(nameof(application));
            CapturedThread = capturedThread;
        }

        public Outlook.Application Application { get; }

        public OutlookThreadContext CapturedThread { get; }

        public OutlookThreadContext AssertAndCaptureCurrentThread()
        {
            var current = OutlookThreadContext.Capture();
            if (current.ManagedThreadId != CapturedThread.ManagedThreadId ||
                current.NativeThreadId != CapturedThread.NativeThreadId ||
                current.ApartmentState != ApartmentState.STA)
            {
                throw new InvalidOperationException(
                    "Outlook work executed outside the captured UI STA.");
            }

            return current;
        }
    }
}
