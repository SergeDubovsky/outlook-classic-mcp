using System;
using OutlookClassicMcp.Core.Outlook;

namespace OutlookClassicMcp.AddIn.Runtime
{
    internal static class OutlookSearchFolderPolicy
    {
        public const int MaximumSearchFoldersExamined = 1024;

        public static void RequireBoundedCount(int count)
        {
            if (count < 0)
            {
                throw new OutlookGatewayException(OutlookGatewayFailure.UnsupportedStore);
            }

            if (count > MaximumSearchFoldersExamined)
            {
                throw new OutlookGatewayException(OutlookGatewayFailure.Timeout);
            }
        }

        public static void RejectIfMatch(bool isSearchFolder)
        {
            if (isSearchFolder)
            {
                throw new OutlookGatewayException(OutlookGatewayFailure.UnsupportedStore);
            }
        }
    }
}
