using System.Threading;
using System.Threading.Tasks;

namespace OutlookClassicMcp.Core.Outlook
{
    public interface IOutlookGateway
    {
        Task<OutlookProbeSnapshot> ProbeAsync(CancellationToken cancellationToken);
    }
}
