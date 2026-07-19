namespace OutlookClassicMcp.Core.Outlook
{
    public sealed class OutlookStoreCapabilities
    {
        public OutlookStoreCapabilities(
            bool isExchangeStore,
            bool isDataFileStore,
            bool isCachedExchange)
        {
            IsExchangeStore = isExchangeStore;
            IsDataFileStore = isDataFileStore;
            IsCachedExchange = isCachedExchange;
        }

        public bool IsExchangeStore { get; }

        public bool IsDataFileStore { get; }

        public bool IsCachedExchange { get; }
    }
}
