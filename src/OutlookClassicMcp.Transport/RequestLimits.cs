namespace OutlookClassicMcp.Transport
{
    public static class RequestLimits
    {
        public static readonly System.TimeSpan DefaultHandlerDeadline =
            System.TimeSpan.FromSeconds(15);

        public const int MaximumHeaderFields = 64;
        public const int MaximumHeaderCharacters = 16 * 1024;
        public const int MaximumHeaderValueCharacters = 8 * 1024;
        public const long MaximumRequestBodyBytes = 1024L * 1024L;
    }
}
