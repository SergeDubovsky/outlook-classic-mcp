namespace OutlookClassicMcp.AddIn.Runtime
{
    internal enum HostLifecycleState
    {
        Created = 0,
        Starting = 1,
        Online = 2,
        Degraded = 3,
        Pausing = 4,
        Paused = 5,
        Stopping = 6,
        Stopped = 7,
    }
}
