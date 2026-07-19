using System;

namespace OutlookClassicMcp.Core.Outlook
{
    public sealed class OutlookDispatcherThreadProof
    {
        public OutlookDispatcherThreadProof(
            int capturedManagedThreadId,
            uint capturedNativeThreadId,
            int executedManagedThreadId,
            uint executedNativeThreadId,
            bool executedOnSta)
        {
            if (capturedManagedThreadId <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(capturedManagedThreadId));
            }

            if (capturedNativeThreadId == 0)
            {
                throw new ArgumentOutOfRangeException(nameof(capturedNativeThreadId));
            }

            if (executedManagedThreadId <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(executedManagedThreadId));
            }

            if (executedNativeThreadId == 0)
            {
                throw new ArgumentOutOfRangeException(nameof(executedNativeThreadId));
            }

            if (capturedManagedThreadId != executedManagedThreadId)
            {
                throw new ArgumentException(
                    "The operation must execute on the captured managed thread.",
                    nameof(executedManagedThreadId));
            }

            if (capturedNativeThreadId != executedNativeThreadId)
            {
                throw new ArgumentException(
                    "The operation must execute on the captured native thread.",
                    nameof(executedNativeThreadId));
            }

            if (!executedOnSta)
            {
                throw new ArgumentException(
                    "The operation must execute in a single-threaded apartment.",
                    nameof(executedOnSta));
            }

            CapturedManagedThreadId = capturedManagedThreadId;
            CapturedNativeThreadId = capturedNativeThreadId;
            ExecutedManagedThreadId = executedManagedThreadId;
            ExecutedNativeThreadId = executedNativeThreadId;
            ExecutedOnSta = executedOnSta;
        }

        public int CapturedManagedThreadId { get; }

        public uint CapturedNativeThreadId { get; }

        public int ExecutedManagedThreadId { get; }

        public uint ExecutedNativeThreadId { get; }

        public bool ExecutedOnSta { get; }
    }
}
