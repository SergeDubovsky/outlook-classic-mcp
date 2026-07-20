using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Windows.Forms;
using OutlookClassicMcp.AddIn.Runtime;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookClassicMcp.AddIn
{
    /// <summary>
    /// Smoke-build-only fixture provisioning that runs entirely inside OUTLOOK.EXE.
    /// </summary>
    internal sealed class Phase4FixtureSeeder : IDisposable
    {
        internal const string ActionEnvironmentVariable = "OUTLOOK_MCP_SMOKE_SEEDER_ACTION";
        internal const string RunIdEnvironmentVariable = "OUTLOOK_MCP_SMOKE_SEEDER_RUN_ID";
        internal const string RunDirectoryEnvironmentVariable = "OUTLOOK_MCP_SMOKE_SEEDER_RUN_DIRECTORY";
        internal const string ExpectedProfileEnvironmentVariable = "OUTLOOK_MCP_SMOKE_SEEDER_EXPECTED_PROFILE";
        internal const string PstAEnvironmentVariable = "OUTLOOK_MCP_SMOKE_SEEDER_PST_A";
        internal const string PstBEnvironmentVariable = "OUTLOOK_MCP_SMOKE_SEEDER_PST_B";
        internal const string StatusPathEnvironmentVariable = "OUTLOOK_MCP_SMOKE_SEEDER_STATUS_PATH";

        private const int BatchSize = 25;
        private const int LargeFolderItemCount = 1001;
        private const int LongBodyMinimumCharacters = 1024;
        private const int MaximumAccountlessProfileStoreCount = 3;
        private const int StaticPaginationItemCount = 12;
        private const string FixtureFileName = "read-fixture.local.json";
        private const string InventoryFileName = "store-inventory.local.json";
        private const string ProgressFileName = "seeder-progress.local.json";
        private const string SourceName = "conditional-vsto-seeder";

        private readonly Outlook.Application _application;
        private readonly System.Windows.Forms.Timer _timer;
        private readonly OutlookThreadContext _ownerThread;
        private readonly string? _requestedAction;
        private SeederConfiguration? _configuration;
        private SeederStage _stage = SeederStage.Validate;
        private bool _disposed;
        private int _detachedStoreCount;
        private int[] _attachmentSizes = Array.Empty<int>();
        private string[] _orderedStaticSubjects = Array.Empty<string>();

        private Phase4FixtureSeeder(Outlook.Application application, string requestedAction)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _requestedAction = requestedAction;
            _ownerThread = OutlookThreadContext.Capture();
            _timer = new System.Windows.Forms.Timer
            {
                Interval = 1,
            };
            _timer.Tick += OnTick;
        }

        /// <summary>
        /// Starts the seeder only when an explicit smoke action is present. A true return value
        /// means the caller must not start the normal MCP host for this Outlook process.
        /// </summary>
        internal static bool TryStart(
            Outlook.Application application,
            out Phase4FixtureSeeder? seeder)
        {
            var action = Environment.GetEnvironmentVariable(ActionEnvironmentVariable);
            if (string.IsNullOrWhiteSpace(action))
            {
                seeder = null;
                return false;
            }

            seeder = new Phase4FixtureSeeder(application, action);
            seeder._timer.Start();
            return true;
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            _timer.Stop();
            _timer.Tick -= OnTick;
            _timer.Dispose();
        }

        private void OnTick(object? sender, EventArgs e)
        {
            if (_disposed)
            {
                return;
            }

            _timer.Stop();
            try
            {
                AssertOwnerThread();
                WriteProgress(_stage, completedCount: 0, targetCount: 0);
                ExecuteStage();
            }
            catch (SeederException exception)
            {
                CompleteFailure(exception.Code);
                return;
            }
            catch (UnauthorizedAccessException)
            {
                CompleteFailure("ACCESS_DENIED");
                return;
            }
            catch (IOException)
            {
                CompleteFailure("IO_FAILURE");
                return;
            }
            catch (COMException)
            {
                CompleteFailure("OUTLOOK_COM_FAILURE");
                return;
            }
            catch (Exception)
            {
                CompleteFailure("INTERNAL_ERROR");
                return;
            }

            if (!_disposed && _stage != SeederStage.Completed)
            {
                _timer.Start();
            }
        }

        private void ExecuteStage()
        {
            if (_configuration != null)
            {
                ValidateConfiguredPaths(_configuration);
            }

            switch (_stage)
            {
                case SeederStage.Validate:
                    _configuration = SeederConfiguration.Load(_requestedAction);
                    WriteProgress(_stage, completedCount: 0, targetCount: 0);
                    ValidateRuntime(_configuration);
                    _stage = _configuration.Action == SeederAction.Seed
                        ? SeederStage.AttachStoreA
                        : SeederStage.DetachStoreA;
                    break;
                case SeederStage.AttachStoreA:
                    EnsureStoreAttached(Configuration.PstAPath, StoreDisplayNameA);
                    _stage = SeederStage.AttachStoreB;
                    break;
                case SeederStage.AttachStoreB:
                    EnsureStoreAttached(Configuration.PstBPath, StoreDisplayNameB);
                    _stage = SeederStage.ValidateDefaultFolders;
                    break;
                case SeederStage.ValidateDefaultFolders:
                    ValidateFixtureDefaultFolders();
                    _stage = SeederStage.SeedKnownBootstrap;
                    break;
                case SeederStage.SeedKnownBootstrap:
                    EnsureBootstrapInboxMessage();
                    _stage = SeederStage.SeedKnownA;
                    break;
                case SeederStage.SeedKnownA:
                    EnsureKnownFolderMessage(
                        Configuration.PstAPath,
                        KnownFolderNameA,
                        KnownSubjectA,
                        KnownBodyA);
                    _stage = SeederStage.SeedKnownB;
                    break;
                case SeederStage.SeedKnownB:
                    EnsureKnownFolderMessage(
                        Configuration.PstBPath,
                        KnownFolderNameB,
                        KnownSubjectB,
                        KnownBodyB);
                    _stage = SeederStage.SeedStaticPagination;
                    break;
                case SeederStage.SeedStaticPagination:
                    if (EnsureStaticPaginationBatch())
                    {
                        _orderedStaticSubjects = ReadStaticOrder();
                        _stage = SeederStage.SeedLargeFolder;
                    }
                    break;
                case SeederStage.SeedLargeFolder:
                    if (EnsureLargeFolderCopyBatch())
                    {
                        _stage = SeederStage.SeedConversationSingleton;
                    }
                    break;
                case SeederStage.SeedConversationSingleton:
                    EnsureSingleFolderMessage(
                        Configuration.PstAPath,
                        ConversationFolderName,
                        ConversationSubject,
                        ConversationBody);
                    _stage = SeederStage.SeedAttachmentMessage;
                    break;
                case SeederStage.SeedAttachmentMessage:
                    _attachmentSizes = EnsureAttachmentMessage();
                    _stage = SeederStage.SeedLongBodyMessage;
                    break;
                case SeederStage.SeedLongBodyMessage:
                    EnsureSingleFolderMessage(
                        Configuration.PstAPath,
                        LongBodyFolderName,
                        LongBodySubject,
                        LongBodyContent);
                    _stage = SeederStage.WriteOutputs;
                    break;
                case SeederStage.WriteOutputs:
                    WriteFixtureOutputs();
                    CompleteSuccess();
                    break;
                case SeederStage.DetachStoreA:
                    if (DetachStore(Configuration.PstAPath))
                    {
                        _detachedStoreCount++;
                    }
                    _stage = SeederStage.DetachStoreB;
                    break;
                case SeederStage.DetachStoreB:
                    if (DetachStore(Configuration.PstBPath))
                    {
                        _detachedStoreCount++;
                    }
                    _stage = SeederStage.VerifyDetached;
                    break;
                case SeederStage.VerifyDetached:
                    VerifyStoresDetached();
                    CompleteSuccess();
                    break;
                case SeederStage.Completed:
                    break;
                default:
                    throw new SeederException("INVALID_STATE");
            }
        }

        private SeederConfiguration Configuration =>
            _configuration ?? throw new SeederException("INVALID_STATE");

        private string RunMarker => Configuration.RunId.ToString("N");

        private string MarkerPrefix => "OCM-P4-" + RunMarker;

        private string BootstrapStoreAlias => "bootstrap_store";

        private string StoreAliasA => "fixture_store_a";

        private string StoreAliasB => "fixture_store_b";

        private string StoreDisplayNameA => MarkerPrefix + "-Store-A";

        private string StoreDisplayNameB => MarkerPrefix + "-Store-B";

        private string StaticFolderName => MarkerPrefix + "-Static";

        private string LargeFolderName => MarkerPrefix + "-Large";

        private string ConversationFolderName => MarkerPrefix + "-Conversation";

        private string AttachmentFolderName => MarkerPrefix + "-Attachments";

        private string LongBodyFolderName => MarkerPrefix + "-LongBody";

        private string KnownFolderNameA => MarkerPrefix + "-Known-A";

        private string KnownFolderNameB => MarkerPrefix + "-Known-B";

        private string BootstrapKnownSubject => MarkerPrefix + "-Bootstrap-Known-Subject";

        private string BootstrapKnownBody => MarkerPrefix + "-Bootstrap-Known-Body";

        private string KnownSubjectA => MarkerPrefix + "-Known-A-Subject";

        private string KnownBodyA => MarkerPrefix + "-Known-A-Body";

        private string KnownSubjectB => MarkerPrefix + "-Known-B-Subject";

        private string KnownBodyB => MarkerPrefix + "-Known-B-Body";

        private string ConversationSeedMarker => MarkerPrefix + "-Conversation-Seed";

        private string ConversationExpectedMarker => MarkerPrefix + "-Conversation-Expected";

        private string ConversationSubject =>
            ConversationSeedMarker + " " + ConversationExpectedMarker;

        private string ConversationBody => MarkerPrefix + "-Conversation-Singleton-Body";

        private string AttachmentSubject => MarkerPrefix + "-Attachment-Subject";

        private string LongBodySubject => MarkerPrefix + "-Long-Body-Subject";

        private string LongBodyPrefix => MarkerPrefix + "-Long-Body-Prefix";

        private string LongBodyContent =>
            LongBodyPrefix + Environment.NewLine + new string('x', LongBodyMinimumCharacters + 256);

        private string AttachmentFileNameA => MarkerPrefix + "-Attachment-A.txt";

        private string AttachmentFileNameB => MarkerPrefix + "-Attachment-B.bin";

        private string AttachmentSourcePathA =>
            Path.Combine(Configuration.RunDirectory, AttachmentFileNameA);

        private string AttachmentSourcePathB =>
            Path.Combine(Configuration.RunDirectory, AttachmentFileNameB);

        private string StaticSubject(int index) =>
            MarkerPrefix + "-Static-" + index.ToString("D2", System.Globalization.CultureInfo.InvariantCulture);

        private string StaticBody(int index) =>
            MarkerPrefix + "-Static-Body-" + index.ToString("D2", System.Globalization.CultureInfo.InvariantCulture);

        private string LargeSharedSubject => MarkerPrefix + "-Large-Shared-Subject";

        private string LargeSharedBody => MarkerPrefix + "-Large-Shared-Body";

        private void AssertOwnerThread()
        {
            var current = OutlookThreadContext.Capture();
            if (_ownerThread.ApartmentState != ApartmentState.STA ||
                current.ManagedThreadId != _ownerThread.ManagedThreadId ||
                current.NativeThreadId != _ownerThread.NativeThreadId ||
                current.ApartmentState != ApartmentState.STA)
            {
                throw new SeederException("WRONG_OUTLOOK_THREAD");
            }
        }

        private void ValidateRuntime(SeederConfiguration configuration)
        {
            AssertOwnerThread();
            ValidateConfiguredPaths(configuration);

            var session = _application.Session;
            if (session == null)
            {
                throw new SeederException("OUTLOOK_SESSION_UNAVAILABLE");
            }

            if (!string.Equals(
                session.CurrentProfileName,
                configuration.ExpectedProfile,
                StringComparison.Ordinal))
            {
                throw new SeederException("PROFILE_MISMATCH");
            }

            Outlook.Accounts? accounts = null;
            try
            {
                accounts = session.Accounts;
                if (accounts == null || accounts.Count != 0)
                {
                    throw new SeederException("PROFILE_HAS_ACCOUNTS");
                }
            }
            finally
            {
                if (accounts != null)
                {
                    Marshal.ReleaseComObject(accounts);
                }
            }

            ValidateProfileStoreIsolation(session, configuration);
            if (configuration.Action == SeederAction.Seed)
            {
                ValidateNoTemplateResidue(session);
            }
        }

        private static void ValidateConfiguredPaths(SeederConfiguration configuration)
        {
            ValidateSecureRunDirectory(configuration.RunDirectory);
            CanonicalizeChildPath(configuration.RunDirectory, configuration.PstAPath, ".pst");
            CanonicalizeChildPath(configuration.RunDirectory, configuration.PstBPath, ".pst");
            CanonicalizeChildPath(configuration.RunDirectory, configuration.StatusPath, ".json");
        }

        private void ValidateNoTemplateResidue(Outlook.NameSpace session)
        {
            Outlook.Inspectors? inspectors = null;
            Outlook.MAPIFolder? drafts = null;
            Outlook.Items? items = null;
            try
            {
                inspectors = _application.Inspectors;
                if (inspectors == null || inspectors.Count != 0)
                {
                    throw new SeederException("FIXTURE_TEMPLATE_RESIDUE");
                }

                drafts = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts)
                    ?? throw new SeederException("BOOTSTRAP_DRAFTS_MISSING");
                items = drafts.Items ?? throw new SeederException("BOOTSTRAP_DRAFTS_MISSING");
                for (var index = 1; index <= items.Count; index++)
                {
                    object? rawItem = null;
                    try
                    {
                        rawItem = items[index];
                        if (rawItem is Outlook.MailItem mail &&
                            IsFixtureTemplateSubject(mail.Subject))
                        {
                            throw new SeederException("FIXTURE_TEMPLATE_RESIDUE");
                        }
                    }
                    finally
                    {
                        if (rawItem != null && Marshal.IsComObject(rawItem))
                        {
                            Marshal.ReleaseComObject(rawItem);
                        }
                    }
                }
            }
            finally
            {
                if (items != null)
                {
                    Marshal.ReleaseComObject(items);
                }
                if (drafts != null)
                {
                    Marshal.ReleaseComObject(drafts);
                }
                if (inspectors != null)
                {
                    Marshal.ReleaseComObject(inspectors);
                }
            }
        }

        private bool IsFixtureTemplateSubject(string? subject)
        {
            if (string.IsNullOrWhiteSpace(subject))
            {
                return false;
            }

            return string.Equals(subject, BootstrapKnownSubject, StringComparison.Ordinal) ||
                string.Equals(subject, KnownSubjectA, StringComparison.Ordinal) ||
                string.Equals(subject, KnownSubjectB, StringComparison.Ordinal) ||
                string.Equals(subject, ConversationSubject, StringComparison.Ordinal) ||
                string.Equals(subject, AttachmentSubject, StringComparison.Ordinal) ||
                string.Equals(subject, LongBodySubject, StringComparison.Ordinal) ||
                string.Equals(subject, LargeSharedSubject, StringComparison.Ordinal) ||
                subject!.StartsWith(MarkerPrefix + "-Static-", StringComparison.Ordinal) ||
                subject.StartsWith(MarkerPrefix + "-Large-", StringComparison.Ordinal);
        }

        private static void ValidateProfileStoreIsolation(
            Outlook.NameSpace session,
            SeederConfiguration configuration)
        {
            Outlook.Stores? stores = null;
            Outlook.Store? defaultStore = null;
            try
            {
                defaultStore = session.DefaultStore;
                stores = session.Stores;
                if (defaultStore == null || stores == null)
                {
                    throw new SeederException("OUTLOOK_SESSION_UNAVAILABLE");
                }

                var bootstrapPath = ValidatePstStore(defaultStore);
                var allowedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    bootstrapPath,
                    configuration.PstAPath,
                    configuration.PstBPath,
                };
                if (allowedPaths.Count != MaximumAccountlessProfileStoreCount ||
                    stores.Count < 1 ||
                    stores.Count > MaximumAccountlessProfileStoreCount)
                {
                    throw new SeederException("PROFILE_HAS_UNEXPECTED_STORES");
                }

                var observedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (var index = 1; index <= stores.Count; index++)
                {
                    Outlook.Store? store = null;
                    try
                    {
                        store = stores[index];
                        if (store == null)
                        {
                            throw new SeederException("PROFILE_HAS_UNEXPECTED_STORES");
                        }

                        var storePath = ValidatePstStore(store);
                        if (!allowedPaths.Contains(storePath) || !observedPaths.Add(storePath))
                        {
                            throw new SeederException("PROFILE_HAS_UNEXPECTED_STORES");
                        }
                    }
                    finally
                    {
                        if (store != null)
                        {
                            Marshal.ReleaseComObject(store);
                        }
                    }
                }

                if (!observedPaths.Contains(bootstrapPath))
                {
                    throw new SeederException("PROFILE_HAS_UNEXPECTED_STORES");
                }
            }
            finally
            {
                if (stores != null)
                {
                    Marshal.ReleaseComObject(stores);
                }
                if (defaultStore != null)
                {
                    Marshal.ReleaseComObject(defaultStore);
                }
            }
        }

        private static string ValidatePstStore(Outlook.Store store)
        {
            var filePath = store.FilePath;
            if (!store.IsDataFileStore ||
                store.ExchangeStoreType != Outlook.OlExchangeStoreType.olNotExchange ||
                string.IsNullOrWhiteSpace(filePath) ||
                !Path.IsPathRooted(filePath) ||
                !string.Equals(Path.GetExtension(filePath), ".pst", StringComparison.OrdinalIgnoreCase))
            {
                throw new SeederException("PROFILE_HAS_UNEXPECTED_STORES");
            }

            try
            {
                var canonical = Path.GetFullPath(filePath);
                if (!string.Equals(canonical, filePath, StringComparison.OrdinalIgnoreCase))
                {
                    throw new SeederException("PROFILE_HAS_UNEXPECTED_STORES");
                }

                return canonical;
            }
            catch (Exception exception) when (
                exception is ArgumentException ||
                exception is NotSupportedException ||
                exception is PathTooLongException)
            {
                throw new SeederException("PROFILE_HAS_UNEXPECTED_STORES");
            }
        }

        private void EnsureStoreAttached(string pstPath, string expectedDisplayName)
        {
            var session = RequireSession();
            Outlook.Store? store = null;
            Outlook.MAPIFolder? root = null;
            try
            {
                store = FindStoreByPath(session, pstPath);
                if (store == null)
                {
                    session.AddStoreEx(pstPath, Outlook.OlStoreType.olStoreUnicode);
                    store = FindStoreByPath(session, pstPath);
                }

                if (store == null ||
                    !PathEquals(store.FilePath, pstPath) ||
                    store.ExchangeStoreType != Outlook.OlExchangeStoreType.olNotExchange)
                {
                    throw new SeederException("PST_ATTACH_FAILED");
                }

                root = store.GetRootFolder()
                    ?? throw new SeederException("PST_ROOT_UNAVAILABLE");
                if (!string.Equals(root.Name, expectedDisplayName, StringComparison.Ordinal))
                {
                    root.Name = expectedDisplayName;
                }
                if (!string.Equals(root.Name, expectedDisplayName, StringComparison.Ordinal))
                {
                    throw new SeederException("PST_DISPLAY_NAME_FAILED");
                }

                Marshal.ReleaseComObject(root);
                root = null;
                Marshal.ReleaseComObject(store);
                store = null;
                store = FindStoreByPath(session, pstPath);
                if (store == null ||
                    !string.Equals(store.DisplayName, expectedDisplayName, StringComparison.Ordinal))
                {
                    throw new SeederException("PST_DISPLAY_NAME_FAILED");
                }

            }
            finally
            {
                if (root != null)
                {
                    Marshal.ReleaseComObject(root);
                }
                if (store != null)
                {
                    Marshal.ReleaseComObject(store);
                }
            }
        }

        private void ValidateFixtureDefaultFolders()
        {
            var session = RequireSession();
            Outlook.Store? store = null;
            Outlook.MAPIFolder? inbox = null;
            try
            {
                store = session.DefaultStore;
                if (store == null ||
                    !store.IsDataFileStore ||
                    store.ExchangeStoreType != Outlook.OlExchangeStoreType.olNotExchange ||
                    PathEquals(store.FilePath, Configuration.PstAPath) ||
                    PathEquals(store.FilePath, Configuration.PstBPath))
                {
                    throw new SeederException("BOOTSTRAP_STORE_INVALID");
                }

                inbox = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                if (inbox == null ||
                    !string.Equals(inbox.StoreID, store.StoreID, StringComparison.Ordinal))
                {
                    throw new SeederException("BOOTSTRAP_INBOX_MISSING");
                }
            }
            catch (COMException)
            {
                throw new SeederException("BOOTSTRAP_INBOX_MISSING");
            }
            finally
            {
                if (inbox != null)
                {
                    Marshal.ReleaseComObject(inbox);
                }
                if (store != null)
                {
                    Marshal.ReleaseComObject(store);
                }
            }
        }

        private void EnsureBootstrapInboxMessage()
        {
            var session = RequireSession();
            Outlook.MAPIFolder? inbox = null;
            try
            {
                inbox = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
                    ?? throw new SeederException("BOOTSTRAP_INBOX_MISSING");
                EnsureMessageExists(
                    inbox,
                    BootstrapKnownSubject,
                    BootstrapKnownBody,
                    requireOnlyItem: false);
            }
            finally
            {
                if (inbox != null)
                {
                    Marshal.ReleaseComObject(inbox);
                }
            }
        }

        private void EnsureKnownFolderMessage(
            string pstPath,
            string folderName,
            string subject,
            string body)
        {
            var session = RequireSession();
            Outlook.Store? store = null;
            Outlook.MAPIFolder? folder = null;
            try
            {
                store = RequireStore(session, pstPath);
                folder = GetOrCreateMailFolder(store, folderName);
                EnsureSinglePersistedFolderMessage(session, store, folder, subject, body);
            }
            finally
            {
                if (folder != null)
                {
                    Marshal.ReleaseComObject(folder);
                }
                if (store != null)
                {
                    Marshal.ReleaseComObject(store);
                }
            }
        }

        private void EnsureSingleFolderMessage(
            string pstPath,
            string folderName,
            string subject,
            string body)
        {
            var session = RequireSession();
            Outlook.Store? store = null;
            Outlook.MAPIFolder? folder = null;
            try
            {
                store = RequireStore(session, pstPath);
                folder = GetOrCreateMailFolder(store, folderName);
                EnsureSinglePersistedFolderMessage(session, store, folder, subject, body);
            }
            finally
            {
                if (folder != null)
                {
                    Marshal.ReleaseComObject(folder);
                }
                if (store != null)
                {
                    Marshal.ReleaseComObject(store);
                }
            }
        }

        private void EnsureSinglePersistedFolderMessage(
            Outlook.NameSpace session,
            Outlook.Store store,
            Outlook.MAPIFolder folder,
            string subject,
            string body)
        {
            Outlook.Items? items = null;
            object? rawItem = null;
            object? rawTemplate = null;
            var templateDiscarded = false;
            try
            {
                var storeId = store.StoreID;
                var folderEntryId = folder.EntryID;
                if (string.IsNullOrWhiteSpace(storeId) || string.IsNullOrWhiteSpace(folderEntryId))
                {
                    throw new SeederException("FIXTURE_COPY_DESTINATION_INVALID");
                }

                items = folder.Items ?? throw new SeederException("PST_FOLDER_UNAVAILABLE");
                if (items.Count > 1)
                {
                    throw new SeederException("FIXTURE_FOLDER_DIRTY");
                }
                if (items.Count == 1)
                {
                    rawItem = items[1];
                    if (!(rawItem is Outlook.MailItem existing))
                    {
                        throw new SeederException("FIXTURE_ITEM_INVALID");
                    }
                    ValidatePersistedMailInFolder(
                        session,
                        existing,
                        storeId,
                        folderEntryId,
                        subject,
                        body);
                    return;
                }

                Marshal.ReleaseComObject(items);
                items = null;
                rawTemplate = _application.CreateItem(Outlook.OlItemType.olMailItem);
                if (!(rawTemplate is Outlook.MailItem template) ||
                    !string.IsNullOrWhiteSpace(template.EntryID))
                {
                    throw new SeederException("FIXTURE_COPY_SOURCE_INVALID");
                }
                var entryId = CreateMovedCopyFromUnsavedTemplate(
                    session,
                    template,
                    folder,
                    storeId,
                    folderEntryId,
                    subject,
                    body);
                if (template.Saved ||
                    !string.IsNullOrWhiteSpace(template.EntryID) ||
                    ReadFolderItemCount(folder) != 1)
                {
                    throw new SeederException("FIXTURE_COPY_COUNT_INVALID");
                }
                ReacquireAndValidateMail(
                    session,
                    entryId,
                    storeId,
                    folderEntryId,
                    subject,
                    body);
                DiscardUnsavedTemplate(template);
                templateDiscarded = true;
            }
            catch (SeederException)
            {
                throw;
            }
            catch (COMException)
            {
                throw new SeederException("FIXTURE_COPY_FAILED");
            }
            finally
            {
                if (!templateDiscarded && rawTemplate is Outlook.MailItem pendingTemplate)
                {
                    TryDiscardUnsavedTemplate(pendingTemplate);
                }
                if (rawTemplate != null && Marshal.IsComObject(rawTemplate))
                {
                    Marshal.ReleaseComObject(rawTemplate);
                }
                if (rawItem != null && Marshal.IsComObject(rawItem))
                {
                    Marshal.ReleaseComObject(rawItem);
                }
                if (items != null)
                {
                    Marshal.ReleaseComObject(items);
                }
            }
        }

        private bool EnsureStaticPaginationBatch()
        {
            var session = RequireSession();
            Outlook.Store? store = null;
            Outlook.MAPIFolder? folder = null;
            object? rawTemplate = null;
            var templateDiscarded = false;
            try
            {
                store = RequireStore(session, Configuration.PstBPath);
                folder = GetOrCreateMailFolder(store, StaticFolderName);
                var storeId = store.StoreID;
                var folderEntryId = folder.EntryID;
                if (string.IsNullOrWhiteSpace(storeId) || string.IsNullOrWhiteSpace(folderEntryId))
                {
                    throw new SeederException("FIXTURE_COPY_DESTINATION_INVALID");
                }

                var expected = new Dictionary<string, string>(
                    StaticPaginationItemCount,
                    StringComparer.Ordinal);
                for (var index = 1; index <= StaticPaginationItemCount; index++)
                {
                    expected.Add(StaticSubject(index), StaticBody(index));
                }

                var observed = ReadAndValidateStaticFolder(
                    session,
                    folder,
                    storeId,
                    folderEntryId,
                    expected);
                if (observed.Count == StaticPaginationItemCount)
                {
                    WriteProgress(_stage, observed.Count, StaticPaginationItemCount);
                    return true;
                }

                rawTemplate = _application.CreateItem(Outlook.OlItemType.olMailItem);
                if (!(rawTemplate is Outlook.MailItem template) ||
                    !string.IsNullOrWhiteSpace(template.EntryID))
                {
                    throw new SeederException("FIXTURE_COPY_SOURCE_INVALID");
                }

                var created = 0;
                for (var index = 1;
                    index <= StaticPaginationItemCount && created < BatchSize;
                    index++)
                {
                    var subject = StaticSubject(index);
                    if (observed.Contains(subject))
                    {
                        continue;
                    }

                    var body = StaticBody(index);
                    var entryId = CreateMovedCopyFromUnsavedTemplate(
                        session,
                        template,
                        folder,
                        storeId,
                        folderEntryId,
                        subject,
                        body);
                    created++;
                    var expectedCount = observed.Count + created;
                    if (ReadFolderItemCount(folder) != expectedCount)
                    {
                        throw new SeederException("FIXTURE_COPY_COUNT_INVALID");
                    }
                    ReacquireAndValidateMail(
                        session,
                        entryId,
                        storeId,
                        folderEntryId,
                        subject,
                        body);
                }

                if (template.Saved || !string.IsNullOrWhiteSpace(template.EntryID))
                {
                    throw new SeederException("FIXTURE_COPY_SOURCE_INVALID");
                }
                DiscardUnsavedTemplate(template);
                templateDiscarded = true;

                observed = ReadAndValidateStaticFolder(
                    session,
                    folder,
                    storeId,
                    folderEntryId,
                    expected);
                WriteProgress(_stage, observed.Count, StaticPaginationItemCount);
                return observed.Count == StaticPaginationItemCount;
            }
            catch (SeederException)
            {
                throw;
            }
            catch (COMException)
            {
                throw new SeederException("FIXTURE_COPY_FAILED");
            }
            finally
            {
                if (!templateDiscarded && rawTemplate is Outlook.MailItem pendingTemplate)
                {
                    TryDiscardUnsavedTemplate(pendingTemplate);
                }
                if (rawTemplate != null && Marshal.IsComObject(rawTemplate))
                {
                    Marshal.ReleaseComObject(rawTemplate);
                }
                if (folder != null)
                {
                    Marshal.ReleaseComObject(folder);
                }
                if (store != null)
                {
                    Marshal.ReleaseComObject(store);
                }
            }
        }

        private bool EnsureLargeFolderCopyBatch()
        {
            var session = RequireSession();
            Outlook.Store? store = null;
            Outlook.MAPIFolder? folder = null;
            object? rawTemplate = null;
            object? rawSource = null;
            var templateDiscarded = false;
            try
            {
                store = RequireStore(session, Configuration.PstBPath);
                folder = GetOrCreateMailFolder(store, LargeFolderName);
                var storeId = store.StoreID;
                var folderEntryId = folder.EntryID;
                if (string.IsNullOrWhiteSpace(storeId) || string.IsNullOrWhiteSpace(folderEntryId))
                {
                    throw new SeederException("FIXTURE_COPY_DESTINATION_INVALID");
                }

                var originalCount = ReadFolderItemCount(folder);
                if (originalCount < 0 || originalCount > LargeFolderItemCount)
                {
                    throw new SeederException("FIXTURE_FOLDER_DIRTY");
                }
                if (originalCount == LargeFolderItemCount)
                {
                    ValidateLargeFolder(session, folder, storeId, folderEntryId);
                    WriteProgress(_stage, originalCount, LargeFolderItemCount);
                    return true;
                }

                var created = 0;
                string sourceEntryId;
                if (originalCount == 0)
                {
                    rawTemplate = _application.CreateItem(Outlook.OlItemType.olMailItem);
                    if (!(rawTemplate is Outlook.MailItem template) ||
                        !string.IsNullOrWhiteSpace(template.EntryID))
                    {
                        throw new SeederException("FIXTURE_COPY_SOURCE_INVALID");
                    }
                    sourceEntryId = CreateMovedCopyFromUnsavedTemplate(
                        session,
                        template,
                        folder,
                        storeId,
                        folderEntryId,
                        LargeSharedSubject,
                        LargeSharedBody);
                    created++;
                    if (ReadFolderItemCount(folder) != 1)
                    {
                        throw new SeederException("FIXTURE_COPY_COUNT_INVALID");
                    }
                    ReacquireAndValidateMail(
                        session,
                        sourceEntryId,
                        storeId,
                        folderEntryId,
                        LargeSharedSubject,
                        LargeSharedBody);
                    if (template.Saved || !string.IsNullOrWhiteSpace(template.EntryID))
                    {
                        throw new SeederException("FIXTURE_COPY_SOURCE_INVALID");
                    }
                    DiscardUnsavedTemplate(template);
                    templateDiscarded = true;
                }
                else
                {
                    sourceEntryId = ReadLargeSourceEntryId(
                        session,
                        folder,
                        storeId,
                        folderEntryId);
                }

                rawSource = session.GetItemFromID(sourceEntryId, storeId);
                if (!(rawSource is Outlook.MailItem source))
                {
                    throw new SeederException("FIXTURE_COPY_SOURCE_INVALID");
                }
                ValidatePersistedMailInFolder(
                    session,
                    source,
                    storeId,
                    folderEntryId,
                    LargeSharedSubject,
                    LargeSharedBody);

                var batchTarget = Math.Min(
                    LargeFolderItemCount,
                    originalCount + BatchSize);
                string? lastCopyEntryId = null;
                var batchEntryIds = new HashSet<string>(StringComparer.Ordinal);
                while (originalCount + created < batchTarget)
                {
                    var validateDestination = lastCopyEntryId == null;
                    lastCopyEntryId = CopyPersistedMail(
                        session,
                        source,
                        sourceEntryId,
                        storeId,
                        folderEntryId,
                        validateDestination);
                    if (!batchEntryIds.Add(lastCopyEntryId))
                    {
                        throw new SeederException("FIXTURE_COPY_RESULT_INVALID");
                    }
                    created++;
                    if (validateDestination)
                    {
                        if (ReadFolderItemCount(folder) != originalCount + created)
                        {
                            throw new SeederException("FIXTURE_COPY_COUNT_INVALID");
                        }
                        ReacquireAndValidateMail(
                            session,
                            lastCopyEntryId,
                            storeId,
                            folderEntryId,
                            LargeSharedSubject,
                            LargeSharedBody);
                    }
                }

                var finalCount = ReadFolderItemCount(folder);
                if (finalCount != originalCount + created)
                {
                    throw new SeederException("FIXTURE_COPY_COUNT_INVALID");
                }
                if (lastCopyEntryId != null)
                {
                    ReacquireAndValidateMail(
                        session,
                        lastCopyEntryId,
                        storeId,
                        folderEntryId,
                        LargeSharedSubject,
                        LargeSharedBody);
                }
                WriteProgress(_stage, finalCount, LargeFolderItemCount);
                if (finalCount != LargeFolderItemCount)
                {
                    return false;
                }

                ValidateLargeFolder(session, folder, storeId, folderEntryId);
                return true;
            }
            catch (SeederException)
            {
                throw;
            }
            catch (COMException)
            {
                throw new SeederException("FIXTURE_COPY_FAILED");
            }
            finally
            {
                if (rawSource != null && Marshal.IsComObject(rawSource))
                {
                    Marshal.ReleaseComObject(rawSource);
                }
                if (!templateDiscarded && rawTemplate is Outlook.MailItem pendingTemplate)
                {
                    TryDiscardUnsavedTemplate(pendingTemplate);
                }
                if (rawTemplate != null && Marshal.IsComObject(rawTemplate))
                {
                    Marshal.ReleaseComObject(rawTemplate);
                }
                if (folder != null)
                {
                    Marshal.ReleaseComObject(folder);
                }
                if (store != null)
                {
                    Marshal.ReleaseComObject(store);
                }
            }
        }

        private static HashSet<string> ReadAndValidateStaticFolder(
            Outlook.NameSpace session,
            Outlook.MAPIFolder folder,
            string storeId,
            string folderEntryId,
            IReadOnlyDictionary<string, string> expected)
        {
            Outlook.Items? items = null;
            var observedSubjects = new HashSet<string>(StringComparer.Ordinal);
            var observedEntryIds = new HashSet<string>(StringComparer.Ordinal);
            try
            {
                items = folder.Items ?? throw new SeederException("PST_FOLDER_UNAVAILABLE");
                if (items.Count < 0 || items.Count > expected.Count)
                {
                    throw new SeederException("FIXTURE_FOLDER_DIRTY");
                }
                for (var index = 1; index <= items.Count; index++)
                {
                    object? rawItem = null;
                    try
                    {
                        rawItem = items[index];
                        if (!(rawItem is Outlook.MailItem mail) ||
                            string.IsNullOrWhiteSpace(mail.Subject) ||
                            !expected.TryGetValue(mail.Subject, out var expectedBody))
                        {
                            throw new SeederException("FIXTURE_ITEM_INVALID");
                        }
                        var entryId = ValidatePersistedMailInFolder(
                            session,
                            mail,
                            storeId,
                            folderEntryId,
                            mail.Subject,
                            expectedBody);
                        if (!observedSubjects.Add(mail.Subject) ||
                            !observedEntryIds.Add(entryId))
                        {
                            throw new SeederException("FIXTURE_ITEM_INVALID");
                        }
                    }
                    finally
                    {
                        if (rawItem != null && Marshal.IsComObject(rawItem))
                        {
                            Marshal.ReleaseComObject(rawItem);
                        }
                    }
                }
                return observedSubjects;
            }
            finally
            {
                if (items != null)
                {
                    Marshal.ReleaseComObject(items);
                }
            }
        }

        private string ReadLargeSourceEntryId(
            Outlook.NameSpace session,
            Outlook.MAPIFolder folder,
            string storeId,
            string folderEntryId)
        {
            Outlook.Items? items = null;
            object? rawItem = null;
            try
            {
                items = folder.Items ?? throw new SeederException("PST_FOLDER_UNAVAILABLE");
                if (items.Count < 1 || items.Count > LargeFolderItemCount)
                {
                    throw new SeederException("FIXTURE_FOLDER_DIRTY");
                }
                rawItem = items[1];
                if (!(rawItem is Outlook.MailItem mail))
                {
                    throw new SeederException("FIXTURE_COPY_SOURCE_INVALID");
                }
                return ValidatePersistedMailInFolder(
                    session,
                    mail,
                    storeId,
                    folderEntryId,
                    LargeSharedSubject,
                    LargeSharedBody);
            }
            finally
            {
                if (rawItem != null && Marshal.IsComObject(rawItem))
                {
                    Marshal.ReleaseComObject(rawItem);
                }
                if (items != null)
                {
                    Marshal.ReleaseComObject(items);
                }
            }
        }

        private void ValidateLargeFolder(
            Outlook.NameSpace session,
            Outlook.MAPIFolder folder,
            string storeId,
            string folderEntryId)
        {
            Outlook.Items? items = null;
            var entryIds = new HashSet<string>(StringComparer.Ordinal);
            try
            {
                items = folder.Items ?? throw new SeederException("PST_FOLDER_UNAVAILABLE");
                if (items.Count != LargeFolderItemCount)
                {
                    throw new SeederException("FIXTURE_FOLDER_DIRTY");
                }
                for (var index = 1; index <= items.Count; index++)
                {
                    object? rawItem = null;
                    try
                    {
                        rawItem = items[index];
                        if (!(rawItem is Outlook.MailItem mail) ||
                            !mail.Saved ||
                            mail.Sent ||
                            string.IsNullOrWhiteSpace(mail.EntryID) ||
                            !entryIds.Add(mail.EntryID))
                        {
                            throw new SeederException("FIXTURE_COPY_RESULT_INVALID");
                        }
                        ValidatePersistedMailInFolder(
                            session,
                            mail,
                            storeId,
                            folderEntryId,
                            LargeSharedSubject,
                            LargeSharedBody);
                    }
                    finally
                    {
                        if (rawItem != null && Marshal.IsComObject(rawItem))
                        {
                            Marshal.ReleaseComObject(rawItem);
                        }
                    }
                }
                if (entryIds.Count != LargeFolderItemCount)
                {
                    throw new SeederException("FIXTURE_COPY_RESULT_INVALID");
                }
            }
            finally
            {
                if (items != null)
                {
                    Marshal.ReleaseComObject(items);
                }
            }
        }

        private static string CreateMovedCopyFromUnsavedTemplate(
            Outlook.NameSpace session,
            Outlook.MailItem template,
            Outlook.MAPIFolder folder,
            string storeId,
            string folderEntryId,
            string subject,
            string body)
        {
            object? rawCopy = null;
            object? rawMoved = null;
            try
            {
                template.Subject = subject;
                template.BodyFormat = Outlook.OlBodyFormat.olFormatPlain;
                template.Body = body;
                template.UnRead = true;
                if (template.Saved || !string.IsNullOrWhiteSpace(template.EntryID))
                {
                    throw new SeederException("FIXTURE_COPY_SOURCE_INVALID");
                }

                rawCopy = template.Copy();
                if (!(rawCopy is Outlook.MailItem copy))
                {
                    throw new SeederException("FIXTURE_COPY_RESULT_INVALID");
                }
                rawMoved = copy.Move(folder);
                if (!(rawMoved is Outlook.MailItem moved))
                {
                    throw new SeederException("FIXTURE_COPY_RESULT_INVALID");
                }
                return ValidatePersistedMailInFolder(
                    session,
                    moved,
                    storeId,
                    folderEntryId,
                    subject,
                    body);
            }
            finally
            {
                if (rawMoved != null && Marshal.IsComObject(rawMoved))
                {
                    Marshal.ReleaseComObject(rawMoved);
                }
                if (rawCopy != null &&
                    !ReferenceEquals(rawCopy, rawMoved) &&
                    Marshal.IsComObject(rawCopy))
                {
                    Marshal.ReleaseComObject(rawCopy);
                }
            }
        }

        private static void DiscardUnsavedTemplate(Outlook.MailItem template)
        {
            try
            {
                template.Close(Outlook.OlInspectorClose.olDiscard);
            }
            catch (COMException)
            {
                throw new SeederException("FIXTURE_TEMPLATE_DISCARD_FAILED");
            }
        }

        private static void TryDiscardUnsavedTemplate(Outlook.MailItem template)
        {
            try
            {
                template.Close(Outlook.OlInspectorClose.olDiscard);
            }
            catch (COMException)
            {
            }
        }

        private string CopyPersistedMail(
            Outlook.NameSpace session,
            Outlook.MailItem source,
            string sourceEntryId,
            string storeId,
            string folderEntryId,
            bool validateDestination)
        {
            object? rawCopy = null;
            try
            {
                rawCopy = source.Copy();
                if (!(rawCopy is Outlook.MailItem copy) ||
                    !copy.Saved ||
                    copy.Sent ||
                    string.IsNullOrWhiteSpace(copy.EntryID) ||
                    !string.Equals(copy.Subject, LargeSharedSubject, StringComparison.Ordinal) ||
                    !string.Equals(copy.Body, LargeSharedBody, StringComparison.Ordinal) ||
                    session.CompareEntryIDs(sourceEntryId, copy.EntryID))
                {
                    throw new SeederException("FIXTURE_COPY_RESULT_INVALID");
                }
                if (validateDestination)
                {
                    ValidatePersistedMailInFolder(
                        session,
                        copy,
                        storeId,
                        folderEntryId,
                        LargeSharedSubject,
                        LargeSharedBody);
                }
                return copy.EntryID;
            }
            finally
            {
                if (rawCopy != null && Marshal.IsComObject(rawCopy))
                {
                    Marshal.ReleaseComObject(rawCopy);
                }
            }
        }

        private static string ValidatePersistedMailInFolder(
            Outlook.NameSpace session,
            Outlook.MailItem mail,
            string storeId,
            string folderEntryId,
            string subject,
            string body)
        {
            if (!mail.Saved ||
                mail.Sent ||
                string.IsNullOrWhiteSpace(mail.EntryID) ||
                !string.Equals(mail.Subject, subject, StringComparison.Ordinal) ||
                !string.Equals(mail.Body, body, StringComparison.Ordinal))
            {
                throw new SeederException("FIXTURE_COPY_RESULT_INVALID");
            }

            object? rawParent = null;
            try
            {
                rawParent = mail.Parent;
                if (!(rawParent is Outlook.MAPIFolder parent) ||
                    string.IsNullOrWhiteSpace(parent.EntryID) ||
                    string.IsNullOrWhiteSpace(parent.StoreID) ||
                    !session.CompareEntryIDs(parent.EntryID, folderEntryId) ||
                    !session.CompareEntryIDs(parent.StoreID, storeId))
                {
                    throw new SeederException("FIXTURE_COPY_DESTINATION_INVALID");
                }
            }
            finally
            {
                if (rawParent != null && Marshal.IsComObject(rawParent))
                {
                    Marshal.ReleaseComObject(rawParent);
                }
            }
            return mail.EntryID;
        }

        private static void ReacquireAndValidateMail(
            Outlook.NameSpace session,
            string entryId,
            string storeId,
            string folderEntryId,
            string subject,
            string body)
        {
            object? rawItem = null;
            try
            {
                rawItem = session.GetItemFromID(entryId, storeId);
                if (!(rawItem is Outlook.MailItem mail) ||
                    string.IsNullOrWhiteSpace(mail.EntryID) ||
                    !session.CompareEntryIDs(mail.EntryID, entryId))
                {
                    throw new SeederException("FIXTURE_COPY_REACQUIRE_FAILED");
                }
                ValidatePersistedMailInFolder(
                    session,
                    mail,
                    storeId,
                    folderEntryId,
                    subject,
                    body);
            }
            finally
            {
                if (rawItem != null && Marshal.IsComObject(rawItem))
                {
                    Marshal.ReleaseComObject(rawItem);
                }
            }
        }

        private static int ReadFolderItemCount(Outlook.MAPIFolder folder)
        {
            Outlook.Items? items = null;
            try
            {
                items = folder.Items ?? throw new SeederException("PST_FOLDER_UNAVAILABLE");
                return items.Count;
            }
            finally
            {
                if (items != null)
                {
                    Marshal.ReleaseComObject(items);
                }
            }
        }

        private int[] EnsureAttachmentMessage()
        {
            EnsureAttachmentSource(AttachmentSourcePathA, 64, 0x41);
            EnsureAttachmentSource(AttachmentSourcePathB, 128, 0x00);

            var session = RequireSession();
            Outlook.Store? store = null;
            Outlook.MAPIFolder? folder = null;
            Outlook.Items? items = null;
            object? rawItem = null;
            Outlook.Attachments? attachments = null;
            try
            {
                store = RequireStore(session, Configuration.PstAPath);
                folder = GetOrCreateMailFolder(store, AttachmentFolderName);
                EnsureSinglePersistedFolderMessage(
                    session,
                    store,
                    folder,
                    AttachmentSubject,
                    MarkerPrefix + "-Attachment-Body");

                items = folder.Items ?? throw new SeederException("PST_FOLDER_UNAVAILABLE");
                if (items.Count != 1)
                {
                    throw new SeederException("FIXTURE_FOLDER_DIRTY");
                }

                rawItem = items[1];
                if (!(rawItem is Outlook.MailItem mail))
                {
                    throw new SeederException("FIXTURE_ITEM_INVALID");
                }

                attachments = mail.Attachments
                    ?? throw new SeederException("FIXTURE_ATTACHMENT_FAILED");
                EnsureAttachment(attachments, AttachmentSourcePathA, AttachmentFileNameA);
                EnsureAttachment(attachments, AttachmentSourcePathB, AttachmentFileNameB);
                mail.Save();

                if (attachments.Count != 2)
                {
                    throw new SeederException("FIXTURE_ATTACHMENT_FAILED");
                }

                var sizes = new int[2];
                for (var index = 1; index <= attachments.Count; index++)
                {
                    Outlook.Attachment? attachment = null;
                    try
                    {
                        attachment = attachments[index];
                        var expectedName = index == 1 ? AttachmentFileNameA : AttachmentFileNameB;
                        if (attachment == null ||
                            !string.Equals(attachment.FileName, expectedName, StringComparison.Ordinal) ||
                            attachment.Size <= 0)
                        {
                            throw new SeederException("FIXTURE_ATTACHMENT_FAILED");
                        }

                        sizes[index - 1] = attachment.Size;
                    }
                    finally
                    {
                        if (attachment != null)
                        {
                            Marshal.ReleaseComObject(attachment);
                        }
                    }
                }

                return sizes;
            }
            finally
            {
                if (attachments != null)
                {
                    Marshal.ReleaseComObject(attachments);
                }
                if (rawItem != null && Marshal.IsComObject(rawItem))
                {
                    Marshal.ReleaseComObject(rawItem);
                }
                if (items != null)
                {
                    Marshal.ReleaseComObject(items);
                }
                if (folder != null)
                {
                    Marshal.ReleaseComObject(folder);
                }
                if (store != null)
                {
                    Marshal.ReleaseComObject(store);
                }
            }
        }

        private static void EnsureAttachment(
            Outlook.Attachments attachments,
            string sourcePath,
            string expectedFileName)
        {
            for (var index = 1; index <= attachments.Count; index++)
            {
                Outlook.Attachment? existing = null;
                try
                {
                    existing = attachments[index];
                    if (existing != null &&
                        string.Equals(existing.FileName, expectedFileName, StringComparison.Ordinal))
                    {
                        return;
                    }
                }
                finally
                {
                    if (existing != null)
                    {
                        Marshal.ReleaseComObject(existing);
                    }
                }
            }

            Outlook.Attachment? added = null;
            try
            {
                added = attachments.Add(
                    sourcePath,
                    Outlook.OlAttachmentType.olByValue,
                    Type.Missing,
                    Type.Missing);
                if (added == null)
                {
                    throw new SeederException("FIXTURE_ATTACHMENT_FAILED");
                }
            }
            finally
            {
                if (added != null)
                {
                    Marshal.ReleaseComObject(added);
                }
            }
        }

        private string[] ReadStaticOrder()
        {
            var session = RequireSession();
            Outlook.Store? store = null;
            Outlook.MAPIFolder? folder = null;
            Outlook.Items? items = null;
            var records = new List<StaticOrderRecord>(StaticPaginationItemCount);
            try
            {
                store = RequireStore(session, Configuration.PstBPath);
                folder = GetExistingMailFolder(store, StaticFolderName);
                items = folder.Items ?? throw new SeederException("PST_FOLDER_UNAVAILABLE");
                if (items.Count != StaticPaginationItemCount)
                {
                    throw new SeederException("FIXTURE_FOLDER_DIRTY");
                }

                for (var index = 1; index <= items.Count; index++)
                {
                    object? rawItem = null;
                    try
                    {
                        rawItem = items[index];
                        if (!(rawItem is Outlook.MailItem mail) ||
                            string.IsNullOrWhiteSpace(mail.EntryID) ||
                            string.IsNullOrWhiteSpace(mail.Subject))
                        {
                            throw new SeederException("FIXTURE_ITEM_INVALID");
                        }

                        records.Add(new StaticOrderRecord(
                            NormalizeOutlookTimestamp(mail.ReceivedTime),
                            mail.EntryID,
                            mail.Subject));
                    }
                    finally
                    {
                        if (rawItem != null && Marshal.IsComObject(rawItem))
                        {
                            Marshal.ReleaseComObject(rawItem);
                        }
                    }
                }
            }
            finally
            {
                if (items != null)
                {
                    Marshal.ReleaseComObject(items);
                }
                if (folder != null)
                {
                    Marshal.ReleaseComObject(folder);
                }
                if (store != null)
                {
                    Marshal.ReleaseComObject(store);
                }
            }

            records.Sort(StaticOrderRecord.Compare);
            return records.Select(record => record.Subject).ToArray();
        }

        private void WriteFixtureOutputs()
        {
            if (_orderedStaticSubjects.Length != StaticPaginationItemCount ||
                _attachmentSizes.Length != 2)
            {
                throw new SeederException("FIXTURE_OUTPUT_INCOMPLETE");
            }

            var bootstrapDisplayName = ReadBootstrapStoreDisplayName();
            var displayNameA = ReadStoreDisplayName(Configuration.PstAPath);
            var displayNameB = ReadStoreDisplayName(Configuration.PstBPath);
            if (string.Equals(bootstrapDisplayName, displayNameA, StringComparison.Ordinal) ||
                string.Equals(bootstrapDisplayName, displayNameB, StringComparison.Ordinal) ||
                string.Equals(displayNameA, displayNameB, StringComparison.Ordinal))
            {
                throw new SeederException("STORE_INVENTORY_FAILED");
            }
            var fixture = new
            {
                schema = 1,
                source = SourceName,
                stores = new object[]
                {
                    new
                    {
                        alias = BootstrapStoreAlias,
                        displayName = bootstrapDisplayName,
                        storeType = "nonExchange",
                        expectsInbox = true,
                        knownMessage = new
                        {
                            subjectMarker = BootstrapKnownSubject,
                            bodyMarker = BootstrapKnownBody,
                        },
                    },
                    new
                    {
                        alias = StoreAliasA,
                        displayName = displayNameA,
                        storeType = "nonExchange",
                        expectsInbox = false,
                        knownMessage = new
                        {
                            folderPath = new[] { KnownFolderNameA },
                            subjectMarker = KnownSubjectA,
                            bodyMarker = KnownBodyA,
                        },
                    },
                    new
                    {
                        alias = StoreAliasB,
                        displayName = displayNameB,
                        storeType = "nonExchange",
                        expectsInbox = false,
                        knownMessage = new
                        {
                            folderPath = new[] { KnownFolderNameB },
                            subjectMarker = KnownSubjectB,
                            bodyMarker = KnownBodyB,
                        },
                    },
                },
                pagination = new
                {
                    storeAlias = StoreAliasB,
                    folderPath = new[] { StaticFolderName },
                    pageSize = 4,
                    orderedSubjectMarkers = _orderedStaticSubjects,
                },
                largeFolder = new
                {
                    storeAlias = StoreAliasB,
                    folderPath = new[] { LargeFolderName },
                    minimumItemCount = LargeFolderItemCount,
                },
                conversation = new
                {
                    storeAlias = StoreAliasA,
                    folderPath = new[] { ConversationFolderName },
                    seedSubjectMarker = ConversationSeedMarker,
                    expectedSubjectMarkers = new[] { ConversationExpectedMarker },
                },
                attachmentMessage = new
                {
                    storeAlias = StoreAliasA,
                    folderPath = new[] { AttachmentFolderName },
                    subjectMarker = AttachmentSubject,
                    expectedAttachments = new[]
                    {
                        new { name = AttachmentFileNameA, size = _attachmentSizes[0] },
                        new { name = AttachmentFileNameB, size = _attachmentSizes[1] },
                    },
                },
                longBodyMessage = new
                {
                    storeAlias = StoreAliasA,
                    folderPath = new[] { LongBodyFolderName },
                    subjectMarker = LongBodySubject,
                    bodyPrefixMarker = LongBodyPrefix,
                    minimumCharacterCount = LongBodyMinimumCharacters,
                },
                protectedMessage = (object?)null,
            };

            var inventory = new
            {
                schema = 1,
                source = SourceName,
                stores = ReadStoreInventory(),
            };

            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = true,
            };
            WriteAtomicSecureJson(
                Path.Combine(Configuration.RunDirectory, FixtureFileName),
                JsonSerializer.Serialize(fixture, options));
            WriteAtomicSecureJson(
                Path.Combine(Configuration.RunDirectory, InventoryFileName),
                JsonSerializer.Serialize(inventory, options));
        }

        private StoreInventoryRecord[] ReadStoreInventory()
        {
            var session = RequireSession();
            Outlook.Stores? stores = null;
            var records = new List<StoreInventoryRecord>();
            try
            {
                stores = session.Stores;
                if (stores == null)
                {
                    throw new SeederException("OUTLOOK_SESSION_UNAVAILABLE");
                }

                for (var index = 1; index <= stores.Count; index++)
                {
                    Outlook.Store? store = null;
                    try
                    {
                        store = stores[index];
                        if (store == null)
                        {
                            throw new SeederException("STORE_INVENTORY_FAILED");
                        }

                        records.Add(new StoreInventoryRecord(
                            store.DisplayName ?? string.Empty,
                            MapStoreType(store.ExchangeStoreType)));
                    }
                    finally
                    {
                        if (store != null)
                        {
                            Marshal.ReleaseComObject(store);
                        }
                    }
                }
            }
            finally
            {
                if (stores != null)
                {
                    Marshal.ReleaseComObject(stores);
                }
            }

            if (records.Count != MaximumAccountlessProfileStoreCount)
            {
                throw new SeederException("STORE_INVENTORY_FAILED");
            }

            return records.ToArray();
        }

        private string ReadBootstrapStoreDisplayName()
        {
            var session = RequireSession();
            Outlook.Store? store = null;
            try
            {
                store = session.DefaultStore;
                if (store == null ||
                    store.ExchangeStoreType != Outlook.OlExchangeStoreType.olNotExchange ||
                    string.IsNullOrWhiteSpace(store.DisplayName))
                {
                    throw new SeederException("BOOTSTRAP_STORE_INVALID");
                }

                return store.DisplayName;
            }
            finally
            {
                if (store != null)
                {
                    Marshal.ReleaseComObject(store);
                }
            }
        }

        private string ReadStoreDisplayName(string pstPath)
        {
            var session = RequireSession();
            Outlook.Store? store = null;
            try
            {
                store = RequireStore(session, pstPath);
                return store.DisplayName ?? string.Empty;
            }
            finally
            {
                if (store != null)
                {
                    Marshal.ReleaseComObject(store);
                }
            }
        }

        private bool DetachStore(string pstPath)
        {
            var session = RequireSession();
            Outlook.Store? store = null;
            Outlook.MAPIFolder? root = null;
            try
            {
                store = FindStoreByPath(session, pstPath);
                if (store == null)
                {
                    return false;
                }

                root = store.GetRootFolder()
                    ?? throw new SeederException("PST_DETACH_FAILED");
                session.RemoveStore(root);
                return true;
            }
            catch (COMException)
            {
                throw new SeederException("PST_DETACH_FAILED");
            }
            finally
            {
                if (root != null)
                {
                    Marshal.ReleaseComObject(root);
                }
                if (store != null)
                {
                    Marshal.ReleaseComObject(store);
                }
            }
        }

        private void VerifyStoresDetached()
        {
            var session = RequireSession();
            Outlook.Store? storeA = null;
            Outlook.Store? storeB = null;
            try
            {
                storeA = FindStoreByPath(session, Configuration.PstAPath);
                storeB = FindStoreByPath(session, Configuration.PstBPath);
                if (storeA != null || storeB != null)
                {
                    throw new SeederException("PST_DETACH_FAILED");
                }
            }
            finally
            {
                if (storeB != null)
                {
                    Marshal.ReleaseComObject(storeB);
                }
                if (storeA != null)
                {
                    Marshal.ReleaseComObject(storeA);
                }
            }
        }

        private Outlook.NameSpace RequireSession()
        {
            AssertOwnerThread();
            return _application.Session
                ?? throw new SeederException("OUTLOOK_SESSION_UNAVAILABLE");
        }

        private static Outlook.Store RequireStore(Outlook.NameSpace session, string pstPath)
        {
            return FindStoreByPath(session, pstPath)
                ?? throw new SeederException("FIXTURE_STORE_NOT_FOUND");
        }

        private static Outlook.Store? FindStoreByPath(Outlook.NameSpace session, string pstPath)
        {
            Outlook.Stores? stores = null;
            try
            {
                stores = session.Stores;
                if (stores == null)
                {
                    return null;
                }

                for (var index = 1; index <= stores.Count; index++)
                {
                    Outlook.Store? candidate = null;
                    try
                    {
                        candidate = stores[index];
                        if (candidate != null && PathEquals(candidate.FilePath, pstPath))
                        {
                            var result = candidate;
                            candidate = null;
                            return result;
                        }
                    }
                    finally
                    {
                        if (candidate != null)
                        {
                            Marshal.ReleaseComObject(candidate);
                        }
                    }
                }

                return null;
            }
            finally
            {
                if (stores != null)
                {
                    Marshal.ReleaseComObject(stores);
                }
            }
        }

        private static Outlook.MAPIFolder GetOrCreateMailFolder(
            Outlook.Store store,
            string folderName)
        {
            Outlook.MAPIFolder? root = null;
            Outlook.Folders? folders = null;
            try
            {
                root = store.GetRootFolder()
                    ?? throw new SeederException("PST_ROOT_UNAVAILABLE");
                folders = root.Folders
                    ?? throw new SeederException("PST_FOLDER_UNAVAILABLE");

                for (var index = 1; index <= folders.Count; index++)
                {
                    Outlook.MAPIFolder? candidate = null;
                    try
                    {
                        candidate = folders[index];
                        if (candidate != null &&
                            string.Equals(candidate.Name, folderName, StringComparison.Ordinal))
                        {
                            var result = candidate;
                            candidate = null;
                            return result;
                        }
                    }
                    finally
                    {
                        if (candidate != null)
                        {
                            Marshal.ReleaseComObject(candidate);
                        }
                    }
                }

                return folders.Add(folderName, Outlook.OlDefaultFolders.olFolderInbox)
                    ?? throw new SeederException("PST_FOLDER_CREATE_FAILED");
            }
            finally
            {
                if (folders != null)
                {
                    Marshal.ReleaseComObject(folders);
                }
                if (root != null)
                {
                    Marshal.ReleaseComObject(root);
                }
            }
        }

        private static Outlook.MAPIFolder GetExistingMailFolder(
            Outlook.Store store,
            string folderName)
        {
            Outlook.MAPIFolder? root = null;
            Outlook.Folders? folders = null;
            try
            {
                root = store.GetRootFolder()
                    ?? throw new SeederException("PST_ROOT_UNAVAILABLE");
                folders = root.Folders
                    ?? throw new SeederException("PST_FOLDER_UNAVAILABLE");
                for (var index = 1; index <= folders.Count; index++)
                {
                    Outlook.MAPIFolder? candidate = null;
                    try
                    {
                        candidate = folders[index];
                        if (candidate != null &&
                            string.Equals(candidate.Name, folderName, StringComparison.Ordinal))
                        {
                            var result = candidate;
                            candidate = null;
                            return result;
                        }
                    }
                    finally
                    {
                        if (candidate != null)
                        {
                            Marshal.ReleaseComObject(candidate);
                        }
                    }
                }
            }
            finally
            {
                if (folders != null)
                {
                    Marshal.ReleaseComObject(folders);
                }
                if (root != null)
                {
                    Marshal.ReleaseComObject(root);
                }
            }

            throw new SeederException("PST_FOLDER_UNAVAILABLE");
        }

        private static void EnsureMessageExists(
            Outlook.MAPIFolder folder,
            string subject,
            string body,
            bool requireOnlyItem)
        {
            Outlook.Items? items = null;
            object? found = null;
            try
            {
                items = folder.Items ?? throw new SeederException("PST_FOLDER_UNAVAILABLE");
                if (requireOnlyItem && items.Count > 1)
                {
                    throw new SeederException("FIXTURE_FOLDER_DIRTY");
                }

                for (var index = 1; index <= items.Count; index++)
                {
                    object? candidate = null;
                    try
                    {
                        candidate = items[index];
                        if (candidate is Outlook.MailItem mail &&
                            string.Equals(mail.Subject, subject, StringComparison.Ordinal))
                        {
                            found = candidate;
                            candidate = null;
                            break;
                        }
                    }
                    finally
                    {
                        if (candidate != null && Marshal.IsComObject(candidate))
                        {
                            Marshal.ReleaseComObject(candidate);
                        }
                    }
                }

                if (found is Outlook.MailItem existing)
                {
                    if (!string.Equals(existing.Body, body, StringComparison.Ordinal))
                    {
                        throw new SeederException("FIXTURE_ITEM_INVALID");
                    }
                    return;
                }

                if (requireOnlyItem && items.Count != 0)
                {
                    throw new SeederException("FIXTURE_FOLDER_DIRTY");
                }

                CreateMailItem(items, subject, body);
            }
            finally
            {
                if (found != null && Marshal.IsComObject(found))
                {
                    Marshal.ReleaseComObject(found);
                }
                if (items != null)
                {
                    Marshal.ReleaseComObject(items);
                }
            }
        }

        private static void CreateMailItem(
            Outlook.Items items,
            string subject,
            string body)
        {
            object? rawItem = null;
            try
            {
                rawItem = items.Add(Outlook.OlItemType.olMailItem);
                if (!(rawItem is Outlook.MailItem mail))
                {
                    throw new SeederException("FIXTURE_ITEM_CREATE_FAILED");
                }

                mail.Subject = subject;
                mail.BodyFormat = Outlook.OlBodyFormat.olFormatPlain;
                mail.Body = body;
                mail.UnRead = true;
                mail.Save();

                if (string.IsNullOrWhiteSpace(mail.EntryID))
                {
                    throw new SeederException("FIXTURE_ITEM_CREATE_FAILED");
                }
            }
            finally
            {
                if (rawItem != null && Marshal.IsComObject(rawItem))
                {
                    Marshal.ReleaseComObject(rawItem);
                }
            }
        }

        private static void EnsureAttachmentSource(string path, int size, byte value)
        {
            var directory = Path.GetDirectoryName(path)
                ?? throw new SeederException("ATTACHMENT_SOURCE_INVALID");
            ValidateSecureRunDirectory(directory);
            path = CanonicalizeChildPath(directory, path, Path.GetExtension(path));
            var content = Enumerable.Repeat(value, size).ToArray();
            if (File.Exists(path))
            {
                ValidateSecureFile(path);
                var existing = File.ReadAllBytes(path);
                if (!existing.SequenceEqual(content))
                {
                    throw new SeederException("ATTACHMENT_SOURCE_INVALID");
                }
                return;
            }

            using (var stream = new FileStream(
                path,
                FileMode.CreateNew,
                FileAccess.Write,
                FileShare.None))
            {
                stream.Write(content, 0, content.Length);
                stream.Flush(flushToDisk: true);
            }
            ApplySecureFileAcl(path);
            ValidateSecureFile(path);
        }

        private void CompleteSuccess()
        {
            WriteProgress(SeederStage.Completed, completedCount: 0, targetCount: 0);
            var status = Configuration.Action == SeederAction.Seed
                ? new
                {
                    schema = 1,
                    runId = RunMarker,
                    action = "seed",
                    success = true,
                    fixtureStoreCount = 3,
                    fixturePstCount = 2,
                    staticPaginationItemCount = StaticPaginationItemCount,
                    largeFolderItemCount = LargeFolderItemCount,
                    conversationMode = "singletonFallback",
                    attachmentCount = 2,
                    fixtureFileName = FixtureFileName,
                    inventoryFileName = InventoryFileName,
                }
                : (object)new
                {
                    schema = 1,
                    runId = RunMarker,
                    action = "detach",
                    success = true,
                    detachedStoreCount = _detachedStoreCount,
                };

            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
            };
            WriteAtomicSecureJson(
                Configuration.StatusPath,
                JsonSerializer.Serialize(status, options));
            _stage = SeederStage.Completed;
        }

        private void CompleteFailure(string code)
        {
            _timer.Stop();
            var failedStage = _stage.ToString();
            _stage = SeederStage.Completed;
            TryWriteFailureStatus(code, failedStage);
        }

        private void WriteProgress(SeederStage stage, int completedCount, int targetCount)
        {
            if (_configuration == null)
            {
                return;
            }

            var progress = new
            {
                schema = 1,
                runId = RunMarker,
                action = Configuration.Action == SeederAction.Seed ? "seed" : "detach",
                stage = stage.ToString(),
                completedCount,
                targetCount,
            };
            try
            {
                WriteAtomicSecureJson(
                    Path.Combine(Configuration.RunDirectory, ProgressFileName),
                    JsonSerializer.Serialize(
                        progress,
                        new JsonSerializerOptions { WriteIndented = true }));
            }
            catch (IOException exception) when (
                exception.HResult == unchecked((int)0x80070020) ||
                exception.HResult == unchecked((int)0x80070021))
            {
                // A progress reader may briefly prevent atomic replacement; the next tick retries.
            }
        }

        private void TryWriteFailureStatus(string code, string failedStage)
        {
            try
            {
                var runDirectory = Environment.GetEnvironmentVariable(RunDirectoryEnvironmentVariable);
                var statusPath = Environment.GetEnvironmentVariable(StatusPathEnvironmentVariable);
                var runId = Environment.GetEnvironmentVariable(RunIdEnvironmentVariable);
                if (string.IsNullOrWhiteSpace(runDirectory) ||
                    string.IsNullOrWhiteSpace(statusPath) ||
                    string.IsNullOrWhiteSpace(runId))
                {
                    return;
                }

                var canonicalDirectory = CanonicalizeExistingDirectory(runDirectory);
                ValidateSecureRunDirectory(canonicalDirectory);
                var canonicalStatus = CanonicalizeChildPath(
                    canonicalDirectory,
                    statusPath,
                    ".json");
                var status = new
                {
                    schema = 1,
                    runId,
                    action = (_requestedAction ?? string.Empty).Trim().ToLowerInvariant(),
                    success = false,
                    errorCode = code,
                    stage = failedStage,
                };
                WriteAtomicSecureJson(
                    canonicalStatus,
                    JsonSerializer.Serialize(
                        status,
                        new JsonSerializerOptions { WriteIndented = true }));
            }
            catch (Exception)
            {
                // Invalid configuration must not redirect diagnostics outside the secure run directory.
            }
        }

        private static void WriteAtomicSecureJson(string path, string json)
        {
            var directory = Path.GetDirectoryName(path)
                ?? throw new SeederException("OUTPUT_PATH_INVALID");
            ValidateSecureRunDirectory(directory);
            path = CanonicalizeChildPath(directory, path, ".json");
            var temporaryPath = Path.Combine(
                directory,
                "." + Path.GetFileName(path) + "." + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                using (var stream = new FileStream(
                    temporaryPath,
                    FileMode.CreateNew,
                    FileAccess.Write,
                    FileShare.None))
                using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
                {
                    writer.Write(json);
                    writer.Write(Environment.NewLine);
                    writer.Flush();
                    stream.Flush(flushToDisk: true);
                }
                ApplySecureFileAcl(temporaryPath);
                ValidateSecureFile(temporaryPath);
                if (File.Exists(path))
                {
                    ValidateSecureFile(path);
                    File.Replace(temporaryPath, path, null, ignoreMetadataErrors: true);
                }
                else
                {
                    File.Move(temporaryPath, path);
                }
                ApplySecureFileAcl(path);
                ValidateSecureFile(path);
            }
            finally
            {
                if (File.Exists(temporaryPath))
                {
                    File.Delete(temporaryPath);
                }
            }
        }

        private static void ValidateSecureRunDirectory(string directory)
        {
            ValidateLocalDirectoryChain(directory);
            var directoryInfo = new DirectoryInfo(directory);
            if (!directoryInfo.Exists ||
                (directoryInfo.Attributes & FileAttributes.ReparsePoint) != 0)
            {
                throw new SeederException("RUN_DIRECTORY_INVALID");
            }

            using (var currentIdentity = WindowsIdentity.GetCurrent())
            {
                var currentSid = currentIdentity.User
                    ?? throw new SeederException("RUN_DIRECTORY_ACL_INVALID");
                var systemSid = new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null);
                var security = directoryInfo.GetAccessControl(
                    AccessControlSections.Access | AccessControlSections.Owner);
                if (!security.AreAccessRulesProtected ||
                    !currentSid.Equals(security.GetOwner(typeof(SecurityIdentifier))))
                {
                    throw new SeederException("RUN_DIRECTORY_ACL_INVALID");
                }

                var currentUserAllowed = false;
                var systemAllowed = false;
                var rules = security.GetAccessRules(
                    includeExplicit: true,
                    includeInherited: true,
                    targetType: typeof(SecurityIdentifier));
                foreach (FileSystemAccessRule rule in rules)
                {
                    if (!(rule.IdentityReference is SecurityIdentifier sid) ||
                        rule.IsInherited ||
                        rule.AccessControlType != AccessControlType.Allow ||
                        (rule.InheritanceFlags &
                            (InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit)) !=
                            (InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit) ||
                        (rule.FileSystemRights & FileSystemRights.FullControl) !=
                            FileSystemRights.FullControl)
                    {
                        throw new SeederException("RUN_DIRECTORY_ACL_INVALID");
                    }

                    if (sid.Equals(currentSid))
                    {
                        currentUserAllowed = true;
                    }
                    else if (sid.Equals(systemSid))
                    {
                        systemAllowed = true;
                    }
                    else
                    {
                        throw new SeederException("RUN_DIRECTORY_ACL_INVALID");
                    }
                }

                if (!currentUserAllowed || !systemAllowed)
                {
                    throw new SeederException("RUN_DIRECTORY_ACL_INVALID");
                }
            }
        }

        private static void ApplySecureFileAcl(string path)
        {
            using (var currentIdentity = WindowsIdentity.GetCurrent())
            {
                var currentSid = currentIdentity.User
                    ?? throw new SeederException("OUTPUT_ACL_FAILED");
                var systemSid = new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null);
                var security = new FileSecurity();
                security.SetOwner(currentSid);
                security.SetAccessRuleProtection(isProtected: true, preserveInheritance: false);
                security.AddAccessRule(new FileSystemAccessRule(
                    currentSid,
                    FileSystemRights.FullControl,
                    AccessControlType.Allow));
                security.AddAccessRule(new FileSystemAccessRule(
                    systemSid,
                    FileSystemRights.FullControl,
                    AccessControlType.Allow));
                File.SetAccessControl(path, security);
            }
        }

        private static void ValidateSecureFile(string path)
        {
            var fileInfo = new FileInfo(path);
            var directory = fileInfo.DirectoryName
                ?? throw new SeederException("OUTPUT_ACL_FAILED");
            ValidateLocalDirectoryChain(directory);
            if (!fileInfo.Exists || (fileInfo.Attributes & FileAttributes.ReparsePoint) != 0)
            {
                throw new SeederException("OUTPUT_ACL_FAILED");
            }

            using (var currentIdentity = WindowsIdentity.GetCurrent())
            {
                var currentSid = currentIdentity.User
                    ?? throw new SeederException("OUTPUT_ACL_FAILED");
                var systemSid = new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null);
                var security = fileInfo.GetAccessControl(
                    AccessControlSections.Access | AccessControlSections.Owner);
                if (!security.AreAccessRulesProtected ||
                    !currentSid.Equals(security.GetOwner(typeof(SecurityIdentifier))))
                {
                    throw new SeederException("OUTPUT_ACL_FAILED");
                }

                var currentUserAllowed = false;
                var systemAllowed = false;
                var rules = security.GetAccessRules(
                    includeExplicit: true,
                    includeInherited: true,
                    targetType: typeof(SecurityIdentifier));
                foreach (FileSystemAccessRule rule in rules)
                {
                    if (!(rule.IdentityReference is SecurityIdentifier sid) ||
                        rule.IsInherited ||
                        rule.AccessControlType != AccessControlType.Allow ||
                        rule.InheritanceFlags != InheritanceFlags.None ||
                        (rule.FileSystemRights & FileSystemRights.FullControl) !=
                            FileSystemRights.FullControl)
                    {
                        throw new SeederException("OUTPUT_ACL_FAILED");
                    }

                    if (sid.Equals(currentSid))
                    {
                        currentUserAllowed = true;
                    }
                    else if (sid.Equals(systemSid))
                    {
                        systemAllowed = true;
                    }
                    else
                    {
                        throw new SeederException("OUTPUT_ACL_FAILED");
                    }
                }

                if (!currentUserAllowed || !systemAllowed)
                {
                    throw new SeederException("OUTPUT_ACL_FAILED");
                }
            }
        }

        private static string MapStoreType(Outlook.OlExchangeStoreType storeType)
        {
            switch (storeType)
            {
                case Outlook.OlExchangeStoreType.olPrimaryExchangeMailbox:
                    return "primaryExchangeMailbox";
                case Outlook.OlExchangeStoreType.olExchangeMailbox:
                    return "exchangeMailbox";
                case Outlook.OlExchangeStoreType.olExchangePublicFolder:
                    return "exchangePublicFolder";
                case Outlook.OlExchangeStoreType.olAdditionalExchangeMailbox:
                    return "additionalExchangeMailbox";
                case Outlook.OlExchangeStoreType.olNotExchange:
                    return "nonExchange";
                default:
                    return "unknown";
            }
        }

        private static DateTime NormalizeOutlookTimestamp(DateTime value)
        {
            if (value.Year < 1900 || value.Year >= 4500)
            {
                return DateTime.SpecifyKind(DateTime.MinValue, DateTimeKind.Utc);
            }

            if (value.Kind == DateTimeKind.Utc)
            {
                return value;
            }

            var local = value.Kind == DateTimeKind.Local
                ? value
                : DateTime.SpecifyKind(value, DateTimeKind.Local);
            return local.ToUniversalTime();
        }

        private static bool PathEquals(string? left, string right)
        {
            if (string.IsNullOrWhiteSpace(left))
            {
                return false;
            }

            try
            {
                return string.Equals(
                    Path.GetFullPath(left),
                    Path.GetFullPath(right),
                    StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception exception) when (
                exception is ArgumentException ||
                exception is NotSupportedException ||
                exception is PathTooLongException)
            {
                return false;
            }
        }

        private static string CanonicalizeExistingDirectory(string value)
        {
            if (string.IsNullOrWhiteSpace(value) ||
                !Path.IsPathRooted(value) ||
                value.StartsWith(@"\\", StringComparison.Ordinal))
            {
                throw new SeederException("RUN_DIRECTORY_INVALID");
            }

            var canonical = Path.GetFullPath(value);
            if (!string.Equals(value, canonical, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(
                    Path.GetPathRoot(canonical)?.TrimEnd(Path.DirectorySeparatorChar),
                    canonical.TrimEnd(Path.DirectorySeparatorChar),
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new SeederException("RUN_DIRECTORY_INVALID");
            }

            canonical = canonical.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            ValidateLocalDirectoryChain(canonical);
            return canonical;
        }

        private static string CanonicalizeChildPath(
            string runDirectory,
            string value,
            string expectedExtension)
        {
            if (string.IsNullOrWhiteSpace(value) || !Path.IsPathRooted(value))
            {
                throw new SeederException("CHILD_PATH_INVALID");
            }

            var canonical = Path.GetFullPath(value);
            var parent = Path.GetDirectoryName(canonical);
            if (!string.Equals(value, canonical, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(parent, runDirectory, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    Path.GetExtension(canonical),
                    expectedExtension,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new SeederException("CHILD_PATH_INVALID");
            }

            if (File.Exists(canonical) &&
                (File.GetAttributes(canonical) & FileAttributes.ReparsePoint) != 0)
            {
                throw new SeederException("CHILD_PATH_INVALID");
            }

            ValidateLocalDirectoryChain(runDirectory);

            return canonical;
        }

        private static void ValidateLocalDirectoryChain(string directory)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(directory) ||
                    directory.StartsWith(@"\\", StringComparison.Ordinal))
                {
                    throw new SeederException("RUN_DIRECTORY_INVALID");
                }

                var canonical = Path.GetFullPath(directory);
                var root = Path.GetPathRoot(canonical);
                if (string.IsNullOrWhiteSpace(root) ||
                    root.Length != 3 ||
                    !char.IsLetter(root[0]) ||
                    root[1] != Path.VolumeSeparatorChar ||
                    (root[2] != Path.DirectorySeparatorChar &&
                        root[2] != Path.AltDirectorySeparatorChar))
                {
                    throw new SeederException("RUN_DIRECTORY_INVALID");
                }

                var drive = new DriveInfo(root);
                if (!drive.IsReady || drive.DriveType != DriveType.Fixed)
                {
                    throw new SeederException("RUN_DIRECTORY_INVALID");
                }

                var current = new DirectoryInfo(canonical);
                while (current != null)
                {
                    if (!current.Exists ||
                        (current.Attributes & FileAttributes.ReparsePoint) != 0)
                    {
                        throw new SeederException("RUN_DIRECTORY_INVALID");
                    }

                    current = current.Parent;
                }
            }
            catch (SeederException)
            {
                throw;
            }
            catch (Exception exception) when (
                exception is ArgumentException ||
                exception is IOException ||
                exception is NotSupportedException ||
                exception is PathTooLongException ||
                exception is UnauthorizedAccessException)
            {
                throw new SeederException("RUN_DIRECTORY_INVALID");
            }
        }

        private enum SeederAction
        {
            Seed,
            Detach,
        }

        private enum SeederStage
        {
            Validate,
            AttachStoreA,
            AttachStoreB,
            ValidateDefaultFolders,
            SeedKnownBootstrap,
            SeedKnownA,
            SeedKnownB,
            SeedStaticPagination,
            SeedLargeFolder,
            SeedConversationSingleton,
            SeedAttachmentMessage,
            SeedLongBodyMessage,
            WriteOutputs,
            DetachStoreA,
            DetachStoreB,
            VerifyDetached,
            Completed,
        }

        private sealed class SeederConfiguration
        {
            private SeederConfiguration(
                SeederAction action,
                Guid runId,
                string runDirectory,
                string expectedProfile,
                string pstAPath,
                string pstBPath,
                string statusPath)
            {
                Action = action;
                RunId = runId;
                RunDirectory = runDirectory;
                ExpectedProfile = expectedProfile;
                PstAPath = pstAPath;
                PstBPath = pstBPath;
                StatusPath = statusPath;
            }

            public SeederAction Action { get; }

            public Guid RunId { get; }

            public string RunDirectory { get; }

            public string ExpectedProfile { get; }

            public string PstAPath { get; }

            public string PstBPath { get; }

            public string StatusPath { get; }

            public static SeederConfiguration Load(string? requestedAction)
            {
                var normalizedAction = (requestedAction ?? string.Empty).Trim().ToLowerInvariant();
                SeederAction action;
                if (normalizedAction == "seed")
                {
                    action = SeederAction.Seed;
                }
                else if (normalizedAction == "detach")
                {
                    action = SeederAction.Detach;
                }
                else
                {
                    throw new SeederException("ACTION_INVALID");
                }

                var runIdValue = RequireEnvironment(RunIdEnvironmentVariable, 32);
                if (!Guid.TryParseExact(runIdValue, "N", out var runId))
                {
                    throw new SeederException("RUN_ID_INVALID");
                }

                var expectedProfile = RequireEnvironment(ExpectedProfileEnvironmentVariable, 128);
                var runDirectory = CanonicalizeExistingDirectory(
                    RequireEnvironment(RunDirectoryEnvironmentVariable, 1024));
                var pstAPath = CanonicalizeChildPath(
                    runDirectory,
                    RequireEnvironment(PstAEnvironmentVariable, 1024),
                    ".pst");
                var pstBPath = CanonicalizeChildPath(
                    runDirectory,
                    RequireEnvironment(PstBEnvironmentVariable, 1024),
                    ".pst");
                var statusPath = CanonicalizeChildPath(
                    runDirectory,
                    RequireEnvironment(StatusPathEnvironmentVariable, 1024),
                    ".json");

                if (string.Equals(pstAPath, pstBPath, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(pstAPath, statusPath, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(pstBPath, statusPath, StringComparison.OrdinalIgnoreCase))
                {
                    throw new SeederException("CHILD_PATH_INVALID");
                }

                return new SeederConfiguration(
                    action,
                    runId,
                    runDirectory,
                    expectedProfile,
                    pstAPath,
                    pstBPath,
                    statusPath);
            }

            private static string RequireEnvironment(string name, int maximumLength)
            {
                var value = Environment.GetEnvironmentVariable(name);
                if (string.IsNullOrWhiteSpace(value) ||
                    value.Length > maximumLength ||
                    value.IndexOf('\0') >= 0)
                {
                    throw new SeederException("CONFIGURATION_INVALID");
                }

                return value;
            }
        }

        private sealed class SeederException : Exception
        {
            public SeederException(string code)
                : base(code)
            {
                Code = code;
            }

            public string Code { get; }
        }

        private sealed class StaticOrderRecord
        {
            public StaticOrderRecord(DateTime timestampUtc, string entryId, string subject)
            {
                TimestampUtc = timestampUtc;
                EntryId = entryId;
                Subject = subject;
            }

            public DateTime TimestampUtc { get; }

            public string EntryId { get; }

            public string Subject { get; }

            public static int Compare(StaticOrderRecord? left, StaticOrderRecord? right)
            {
                if (ReferenceEquals(left, right))
                {
                    return 0;
                }
                if (left == null)
                {
                    return 1;
                }
                if (right == null)
                {
                    return -1;
                }

                var timestamp = right.TimestampUtc.CompareTo(left.TimestampUtc);
                return timestamp != 0
                    ? timestamp
                    : string.CompareOrdinal(left.EntryId, right.EntryId);
            }
        }

        private sealed class StoreInventoryRecord
        {
            public StoreInventoryRecord(string displayName, string storeType)
            {
                DisplayName = displayName;
                StoreType = storeType;
            }

            public string DisplayName { get; }

            public string StoreType { get; }
        }
    }
}
