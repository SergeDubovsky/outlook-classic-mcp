namespace OutlookClassicMcp.Core.Policy
{
    public enum ImplementationPhase
    {
        RepositoryAndToolchain = 0,
        DependencyAndLifecycle = 1,
        AuthenticatedTransport = 2,
        OutlookProbe = 3,
        BoundedReads = 4,
        Drafts = 5,
        ReversibleMutations = 6,
        Send = 7,
        AttachmentsAndOperations = 8,
        PublicRelease = 9,
    }
}
