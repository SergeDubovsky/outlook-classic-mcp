using OutlookClassicMcp.Core.Outlook;

namespace OutlookClassicMcp.AddIn.Runtime
{
    internal static class OutlookComMetrics
    {
        private static readonly object Gate = new object();
        private static long _acquired;
        private static long _outstanding;
        private static long _peak;
        private static long _released;
        private static long _materializedItemHighWater;

        public static void RecordAcquired()
        {
            lock (Gate)
            {
                _acquired++;
                _outstanding++;
                if (_outstanding > _peak)
                {
                    _peak = _outstanding;
                }
            }
        }

        public static void RecordReleased()
        {
            lock (Gate)
            {
                _released++;
                _outstanding--;
            }
        }

        public static void ObserveMaterializedItems(int count)
        {
            if (count < 0)
            {
                return;
            }

            lock (Gate)
            {
                if (count > _materializedItemHighWater)
                {
                    _materializedItemHighWater = count;
                }
            }
        }

        public static OutlookReadDiagnosticsSnapshot Capture()
        {
            lock (Gate)
            {
                return new OutlookReadDiagnosticsSnapshot(
                    _acquired,
                    _released,
                    _outstanding,
                    _peak,
                    _materializedItemHighWater);
            }
        }

        internal static void ResetForTests()
        {
            lock (Gate)
            {
                _acquired = 0;
                _released = 0;
                _outstanding = 0;
                _peak = 0;
                _materializedItemHighWater = 0;
            }
        }
    }
}
