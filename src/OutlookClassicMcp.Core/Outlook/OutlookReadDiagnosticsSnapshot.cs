using System;

namespace OutlookClassicMcp.Core.Outlook
{
    public sealed class OutlookReadDiagnosticsSnapshot
    {
        public OutlookReadDiagnosticsSnapshot(
            long comAcquired,
            long comReleased,
            long comOutstanding,
            long comPeak,
            long materializedItemHighWater)
        {
            ComAcquired = RequireNonnegative(comAcquired, nameof(comAcquired));
            ComReleased = RequireNonnegative(comReleased, nameof(comReleased));
            ComOutstanding = RequireNonnegative(comOutstanding, nameof(comOutstanding));
            ComPeak = RequireNonnegative(comPeak, nameof(comPeak));
            MaterializedItemHighWater = RequireNonnegative(
                materializedItemHighWater,
                nameof(materializedItemHighWater));
            if (ComReleased > ComAcquired ||
                ComOutstanding != ComAcquired - ComReleased)
            {
                throw new ArgumentException("The COM ownership counters are inconsistent.");
            }

            if (ComPeak < ComOutstanding || ComPeak > ComAcquired)
            {
                throw new ArgumentException("The COM ownership peak is inconsistent.");
            }
        }

        public long ComAcquired { get; }

        public long ComReleased { get; }

        public long ComOutstanding { get; }

        public long ComPeak { get; }

        public long MaterializedItemHighWater { get; }

        private static long RequireNonnegative(long value, string parameterName)
        {
            if (value < 0)
            {
                throw new ArgumentOutOfRangeException(parameterName);
            }

            return value;
        }
    }
}
