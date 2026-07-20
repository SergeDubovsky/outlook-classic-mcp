namespace OutlookClassicMcp.Transport
{
    public static class RequestLimits
    {
        public static readonly System.TimeSpan DefaultHandlerDeadline =
            System.TimeSpan.FromSeconds(15);
        public static readonly System.TimeSpan DefaultToolDeadline =
            System.TimeSpan.FromSeconds(14);

        public const int MaximumHeaderFields = 64;
        public const int MaximumHeaderCharacters = 16 * 1024;
        public const int MaximumHeaderValueCharacters = 8 * 1024;
        public const long MaximumRequestBodyBytes = 1024L * 1024L;
        public const int DefaultReadPageSize = 25;
        public const int DefaultMessageBodyCharacters = 50000;
        public const int MaximumToolResultBytes = 1024 * 1024;
    }
}
