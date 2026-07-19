using System;

namespace OutlookClassicMcp.Core.Outlook
{
    public sealed class OutlookStoreProbe
    {
        public const int MaximumDisplayNameLength = 256;

        public OutlookStoreProbe(
            string displayName,
            OutlookStoreType storeType,
            OutlookStoreCapabilities capabilities,
            StandardFolderAvailability standardFolders)
        {
            DisplayName = OutlookContractValidation.RequireBoundedText(
                displayName,
                MaximumDisplayNameLength,
                nameof(displayName));
            OutlookContractValidation.RequireDefinedEnum(storeType, nameof(storeType));
            StoreType = storeType;
            Capabilities = capabilities ?? throw new ArgumentNullException(nameof(capabilities));
            StandardFolders = standardFolders ?? throw new ArgumentNullException(nameof(standardFolders));
        }

        public string DisplayName { get; }

        public OutlookStoreType StoreType { get; }

        public OutlookStoreCapabilities Capabilities { get; }

        public StandardFolderAvailability StandardFolders { get; }
    }
}
