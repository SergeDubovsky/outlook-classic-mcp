using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace OutlookClassicMcp.AddIn.Runtime
{
    internal readonly struct OutlookThreadContext
    {
        private OutlookThreadContext(int managedThreadId, uint nativeThreadId, ApartmentState apartmentState)
        {
            ManagedThreadId = managedThreadId;
            NativeThreadId = nativeThreadId;
            ApartmentState = apartmentState;
        }

        public int ManagedThreadId { get; }

        public uint NativeThreadId { get; }

        public ApartmentState ApartmentState { get; }

        public static OutlookThreadContext Capture()
        {
            return new OutlookThreadContext(
                Environment.CurrentManagedThreadId,
                NativeMethods.GetCurrentThreadId(),
                Thread.CurrentThread.GetApartmentState());
        }

        private static class NativeMethods
        {
            [DllImport("kernel32.dll")]
            [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
            internal static extern uint GetCurrentThreadId();
        }
    }
}
