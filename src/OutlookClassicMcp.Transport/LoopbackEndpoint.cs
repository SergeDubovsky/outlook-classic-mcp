using System;

namespace OutlookClassicMcp.Transport
{
    public static class LoopbackEndpoint
    {
        public const string Prefix = "http://127.0.0.1:8765/mcp/";

        public static Uri Address { get; } = new Uri(Prefix, UriKind.Absolute);
    }
}
