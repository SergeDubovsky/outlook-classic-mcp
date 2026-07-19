using System;
using System.Collections.Generic;

namespace OutlookClassicMcp.Core.Outlook
{
    public sealed class OutlookProbeSnapshot
    {
        public const int MaximumStoreCount = 64;
        public const int MaximumVersionLength = 64;
        public const int MaximumProfileNameLength = 256;
        public const int MaximumWarningCount = 8;

        public OutlookProbeSnapshot(
            string outlookVersion,
            int outlookBitness,
            string activeProfileName,
            OutlookDispatcherThreadProof dispatcherThread,
            int configuredStoreCount,
            IEnumerable<OutlookStoreProbe> stores)
            : this(
                outlookVersion,
                outlookBitness,
                activeProfileName,
                dispatcherThread,
                configuredStoreCount,
                stores,
                Array.Empty<OutlookProbeWarning>())
        {
        }

        public OutlookProbeSnapshot(
            string outlookVersion,
            int outlookBitness,
            string activeProfileName,
            OutlookDispatcherThreadProof dispatcherThread,
            int configuredStoreCount,
            IEnumerable<OutlookStoreProbe> stores,
            IEnumerable<OutlookProbeWarning> warnings)
        {
            OutlookVersion = OutlookContractValidation.RequireBoundedText(
                outlookVersion,
                MaximumVersionLength,
                nameof(outlookVersion));
            if (outlookBitness != 32 && outlookBitness != 64)
            {
                throw new ArgumentOutOfRangeException(nameof(outlookBitness));
            }

            OutlookBitness = outlookBitness;
            ActiveProfileName = OutlookContractValidation.RequireBoundedText(
                activeProfileName,
                MaximumProfileNameLength,
                nameof(activeProfileName));
            DispatcherThread = dispatcherThread ?? throw new ArgumentNullException(nameof(dispatcherThread));

            if (configuredStoreCount < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(configuredStoreCount));
            }

            var storeCopy = OutlookContractValidation.BoundedCopy(
                stores,
                MaximumStoreCount,
                nameof(stores),
                rejectNullElements: true);
            if (configuredStoreCount < storeCopy.Count)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(configuredStoreCount),
                    "The configured store count must include every returned store.");
            }

            var warningCopy = OutlookContractValidation.BoundedCopy(
                warnings,
                MaximumWarningCount,
                nameof(warnings),
                rejectNullElements: false);
            for (var index = 0; index < warningCopy.Count; index++)
            {
                OutlookContractValidation.RequireDefinedEnum(warningCopy[index], nameof(warnings));
            }

            ConfiguredStoreCount = configuredStoreCount;
            Stores = storeCopy;
            Warnings = warningCopy;
            IsPartial = configuredStoreCount > storeCopy.Count || warningCopy.Count > 0;
        }

        public string OutlookVersion { get; }

        public int OutlookBitness { get; }

        public string ActiveProfileName { get; }

        public OutlookDispatcherThreadProof DispatcherThread { get; }

        public int ConfiguredStoreCount { get; }

        public IReadOnlyList<OutlookStoreProbe> Stores { get; }

        public IReadOnlyList<OutlookProbeWarning> Warnings { get; }

        public bool IsPartial { get; }
    }
}
