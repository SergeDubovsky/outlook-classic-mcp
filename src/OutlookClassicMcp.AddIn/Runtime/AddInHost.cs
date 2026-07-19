using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using OutlookClassicMcp.Transport;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookClassicMcp.AddIn.Runtime
{
    internal sealed class AddInHost
    {
        private readonly Outlook.Application _application;
        private readonly Outlook.ApplicationEvents_11_Event _applicationEvents;
        private readonly MetadataDiagnostics _diagnostics;
        private readonly HostLifecycle _lifecycle = new HostLifecycle();
        private readonly object _shutdownGate = new object();
        private readonly object _transportGate = new object();
        private readonly TaskCompletionSource<bool> _shutdownCompletion =
            new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        private OutlookStaDispatcher? _dispatcher;
        private OutlookGateway? _gateway;
        private LoopbackHttpServer? _server;
        private Task? _initializationTask;
        private Timer? _shutdownWatchdog;
        private Exception? _shutdownFailure;
        private Exception? _startupFailure;
        private long _shutdownStartedTicks;
        private int _shutdownCompletionOwner;
        private int _staCleanupStarted;
        private int _stopRequested;
        private int _trackedTaskCount;
        private bool _quitSubscribed;

        public AddInHost(Outlook.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _applicationEvents = (Outlook.ApplicationEvents_11_Event)application;
            _diagnostics = MetadataDiagnostics.Create();
        }

        public HostLifecycleState State => _lifecycle.State;

        public Task InitializationCompletion => _initializationTask ?? Task.CompletedTask;

        public Task ShutdownCompletion => _shutdownCompletion.Task;

        public void Start(long callbackStartedTicks)
        {
            var priorInitialization = _initializationTask;
            if (priorInitialization != null && !priorInitialization.IsCompleted)
            {
                return;
            }

            if (!_lifecycle.TryBeginStartup())
            {
                return;
            }

            _startupFailure = null;
            var startupMeasurementScheduled = false;
            var startupDurationTicks = 0L;
            try
            {
                var dispatcher = _dispatcher;
                if (dispatcher == null)
                {
                    var runtimeContext = new OutlookRuntimeContext(
                        _application,
                        OutlookThreadContext.Capture());
                    dispatcher = new OutlookStaDispatcher(
                        runtimeContext,
                        OnDispatcherUnavailable);
                    _dispatcher = dispatcher;
                    _gateway = new OutlookGateway(dispatcher);
                }
                else if (!dispatcher.IsAccepting)
                {
                    throw new InvalidOperationException(
                        "The Outlook dispatcher is unavailable. Restart Outlook to recover it.");
                }
                else if (_gateway == null)
                {
                    _gateway = new OutlookGateway(dispatcher);
                }

                dispatcher.AssertOutlookThread();
                if (!_quitSubscribed)
                {
                    _applicationEvents.Quit += OnApplicationQuit;
                    _quitSubscribed = true;
                }

                Interlocked.Increment(ref _trackedTaskCount);
                try
                {
                    startupMeasurementScheduled = dispatcher.TryPostLifecycleCallback(
                        () => RecordStartupCallback(startupDurationTicks));
                    if (!startupMeasurementScheduled)
                    {
                        throw new InvalidOperationException("The startup completion measurement could not be posted.");
                    }

                    _initializationTask = Task.Run((Action)Initialize);
                    startupDurationTicks = Stopwatch.GetTimestamp() - callbackStartedTicks;
                }
                catch
                {
                    Interlocked.Decrement(ref _trackedTaskCount);
                    throw;
                }
            }
            catch (Exception exception)
            {
                startupDurationTicks = Stopwatch.GetTimestamp() - callbackStartedTicks;
                _startupFailure = exception;
                _lifecycle.TryMarkDegraded();
                if (!startupMeasurementScheduled)
                {
                    RecordStartupCallback(Stopwatch.GetTimestamp() - callbackStartedTicks);
                }
            }
        }

        public Task BeginShutdown()
        {
            if (!_lifecycle.TryBeginStop())
            {
                return _shutdownCompletion.Task;
            }

            Interlocked.Exchange(ref _stopRequested, 1);
            _shutdownStartedTicks = Stopwatch.GetTimestamp();
            if (_quitSubscribed)
            {
                try
                {
                    _applicationEvents.Quit -= OnApplicationQuit;
                    _quitSubscribed = false;
                }
                catch (Exception exception)
                {
                    RecordShutdownFailure(exception);
                }
            }

            var dispatcher = _dispatcher;
            if (dispatcher != null)
            {
                try
                {
                    dispatcher.BeginShutdown();
                }
                catch (Exception exception)
                {
                    RecordShutdownFailure(exception);
                }
            }

            BeginListenerShutdown();

            var initializationTask = _initializationTask;
            if (initializationTask == null || initializationTask.IsCompleted)
            {
                ObserveInitialization(initializationTask);
                CompleteShutdownAfterTransport(onOutlookThread: true);
            }
            else
            {
                initializationTask.ConfigureAwait(false).GetAwaiter().OnCompleted(
                    () => CompleteShutdownAfterInitialization(initializationTask));
            }

            return _shutdownCompletion.Task;
        }

        private void RecordStartupCallback(long durationTicks)
        {
            var dispatcher = _dispatcher;
            var queueDepth = dispatcher?.QueueDepth ?? 0;
            var trackedTaskCount = Volatile.Read(ref _trackedTaskCount);
            var failure = _startupFailure;
            if (failure == null)
            {
                _diagnostics.RecordStartup(
                    _lifecycle.State,
                    durationTicks,
                    Debugger.IsAttached,
                    queueDepth,
                    trackedTaskCount,
                    IsListenerActive());
            }
            else
            {
                _diagnostics.RecordFailure(
                    RuntimeDiagnosticEvent.StartupCompleted,
                    _lifecycle.State,
                    durationTicks,
                    queueDepth,
                    trackedTaskCount,
                    failure,
                    IsListenerActive());
            }
        }

        private void Initialize()
        {
            var stage = RuntimeDiagnosticEvent.DependencyBindingCompleted;
            var stageStarted = Stopwatch.GetTimestamp();
            try
            {
                var dependencyReport = McpDependencyProbe.VerifyLoad();
                _diagnostics.RecordDependencyBinding(
                    _lifecycle.State,
                    Stopwatch.GetTimestamp() - stageStarted,
                    dependencyReport.CoreVersion,
                    dependencyReport.LoadedAssemblyIdentities.Count,
                    dependencyReport.IdentitySha256,
                    _dispatcher?.QueueDepth ?? 0,
                    Volatile.Read(ref _trackedTaskCount));

                stage = RuntimeDiagnosticEvent.DispatcherProbeCompleted;
                stageStarted = Stopwatch.GetTimestamp();
                var dispatcher = _dispatcher
                    ?? throw new InvalidOperationException("The Outlook dispatcher is unavailable.");
                var executedThread = dispatcher
                    .InvokeAsync(OutlookThreadContext.Capture, CancellationToken.None)
                    .GetAwaiter()
                    .GetResult();

                var capturedThread = dispatcher.OwnerThread;
                if (capturedThread.ManagedThreadId != executedThread.ManagedThreadId ||
                    capturedThread.NativeThreadId != executedThread.NativeThreadId ||
                    executedThread.ApartmentState != ApartmentState.STA)
                {
                    throw new InvalidOperationException("The dispatcher probe did not return to the captured Outlook STA.");
                }

                _diagnostics.RecordDispatcherProbe(
                    _lifecycle.State,
                    Stopwatch.GetTimestamp() - stageStarted,
                    capturedThread,
                    executedThread,
                    dispatcher.QueueDepth,
                    Volatile.Read(ref _trackedTaskCount));

                stage = RuntimeDiagnosticEvent.ListenerBindingCompleted;
                stageStarted = Stopwatch.GetTimestamp();
                BearerToken? token = null;
                LoopbackHttpServer? candidate = null;
                LoopbackHttpServer? startedServer = null;
                try
                {
                    token = BearerToken.LoadFromProcessEnvironment();
                    var gateway = _gateway
                        ?? throw new InvalidOperationException("The Outlook gateway is unavailable.");
                    candidate = new LoopbackHttpServer(token, CreateStatusSnapshot, gateway);
                    token = null;

                    if (Volatile.Read(ref _stopRequested) != 0)
                    {
                        throw new OperationCanceledException("Shutdown began before listener startup.");
                    }

                    candidate.Start();
                    lock (_transportGate)
                    {
                        if (Volatile.Read(ref _stopRequested) == 0)
                        {
                            _server = candidate;
                            startedServer = candidate;
                            candidate = null;
                        }
                    }

                    if (candidate != null)
                    {
                        throw new OperationCanceledException("Shutdown began during listener startup.");
                    }
                }
                finally
                {
                    candidate?.Dispose();
                    token?.Dispose();
                }

                if (startedServer == null || !startedServer.IsListening)
                {
                    if (startedServer != null)
                    {
                        _lifecycle.TryMarkDegraded();
                        ObserveListenerCompletion(startedServer);
                    }

                    throw new InvalidOperationException("The loopback listener stopped during startup.");
                }

                if (!_lifecycle.TryMarkOnline())
                {
                    BeginListenerShutdown();
                    if (Volatile.Read(ref _stopRequested) != 0)
                    {
                        return;
                    }

                    throw new InvalidOperationException("The host could not transition online after listener startup.");
                }

                if (!startedServer.IsListening)
                {
                    _lifecycle.TryMarkDegraded();
                    ObserveListenerCompletion(startedServer);
                    throw new InvalidOperationException("The loopback listener stopped during startup.");
                }

                _diagnostics.RecordListenerBinding(
                    HostLifecycleState.Online,
                    Stopwatch.GetTimestamp() - stageStarted,
                    dispatcher.QueueDepth,
                    Volatile.Read(ref _trackedTaskCount),
                    listenerActive: true);
                ObserveListenerCompletion(startedServer);
            }
            catch (Exception exception)
            {
                _lifecycle.TryMarkDegraded();
                _diagnostics.RecordFailure(
                    stage,
                    _lifecycle.State,
                    Stopwatch.GetTimestamp() - stageStarted,
                    _dispatcher?.QueueDepth ?? 0,
                    Volatile.Read(ref _trackedTaskCount),
                    exception,
                    IsListenerActive());
            }
            finally
            {
                var remainingTasks = Interlocked.Decrement(ref _trackedTaskCount);
                _diagnostics.RecordHostQuiescent(
                    _lifecycle.State,
                    _dispatcher?.QueueDepth ?? 0,
                    remainingTasks,
                    IsListenerActive());
            }
        }

        private void OnApplicationQuit()
        {
            _ = BeginShutdown();
        }

        private void OnDispatcherUnavailable()
        {
            if (Volatile.Read(ref _stopRequested) != 0)
            {
                return;
            }

            _lifecycle.TryMarkDegraded();
        }

        private void CompleteShutdownAfterInitialization(Task initializationTask)
        {
            ObserveInitialization(initializationTask);
            BeginListenerShutdown();
            CompleteShutdownAfterTransport(onOutlookThread: false);
        }

        private void CompleteShutdownAfterTransport(bool onOutlookThread)
        {
            LoopbackHttpServer? server;
            lock (_transportGate)
            {
                server = _server;
            }

            var completion = server?.BeginShutdown() ?? Task.CompletedTask;

            if (!completion.IsCompleted)
            {
                completion.ConfigureAwait(false).GetAwaiter().OnCompleted(
                    () => CompleteShutdownAfterTransport(onOutlookThread: false));
                return;
            }

            ObserveTransport(completion);
            CompleteShutdownAfterDispatcher(onOutlookThread);
        }

        private void CompleteShutdownAfterDispatcher(bool onOutlookThread)
        {
            var dispatcher = _dispatcher;
            Task completion;
            try
            {
                completion = dispatcher?.BeginShutdown() ?? Task.CompletedTask;
            }
            catch (Exception exception)
            {
                RecordShutdownFailure(exception);
                completion = Task.CompletedTask;
            }

            if (!completion.IsCompleted)
            {
                completion.ConfigureAwait(false).GetAwaiter().OnCompleted(
                    () => CompleteShutdownAfterDispatcher(onOutlookThread: false));
                return;
            }

            try
            {
                completion.GetAwaiter().GetResult();
            }
            catch (Exception exception)
            {
                RecordShutdownFailure(exception);
            }

            if (onOutlookThread)
            {
                FinalizeShutdownOnOutlookThread();
                return;
            }

            if (dispatcher != null && dispatcher.TryPostLifecycleCallback(FinalizeShutdownOnOutlookThread))
            {
                ArmShutdownWatchdog();
                return;
            }

            RecordShutdownFailure(new InvalidOperationException("The Outlook STA was unavailable for final shutdown."));
            CompleteShutdownWithoutOutlookThread();
        }

        private void FinalizeShutdownOnOutlookThread()
        {
            if (Interlocked.Exchange(ref _staCleanupStarted, 1) != 0)
            {
                return;
            }

            var ownsCompletion = Interlocked.CompareExchange(ref _shutdownCompletionOwner, 1, 0) == 0;
            var queueDepth = 0;
            var trackedTaskCount = Volatile.Read(ref _trackedTaskCount);
            var completionRecorded = false;
            try
            {
                DisposeShutdownWatchdog();
                DisposeListener();
                var dispatcher = _dispatcher;
                if (dispatcher != null)
                {
                    queueDepth = dispatcher.QueueDepth;
                    try
                    {
                        dispatcher.Dispose();
                    }
                    catch (Exception exception)
                    {
                        RecordShutdownFailure(exception);
                    }
                    finally
                    {
                        _dispatcher = null;
                        _gateway = null;
                    }
                }

                trackedTaskCount = Volatile.Read(ref _trackedTaskCount);
                if (queueDepth != 0 || trackedTaskCount != 0)
                {
                    RecordShutdownFailure(new InvalidOperationException("HOST_NOT_QUIESCENT"));
                }

                try
                {
                    _lifecycle.MarkStopped();
                }
                catch (Exception exception)
                {
                    RecordShutdownFailure(exception);
                }
                if (ownsCompletion)
                {
                    RecordShutdownCompletion(queueDepth, trackedTaskCount);
                    completionRecorded = true;
                }
            }
            catch (Exception exception)
            {
                RecordShutdownFailure(exception);
            }
            finally
            {
                if (ownsCompletion && !completionRecorded)
                {
                    try
                    {
                        _lifecycle.MarkStopped();
                    }
                    catch (Exception exception)
                    {
                        RecordShutdownFailure(exception);
                    }
                    RecordShutdownCompletion(queueDepth, trackedTaskCount);
                }
            }
        }

        private void CompleteShutdownWithoutOutlookThread()
        {
            if (Interlocked.CompareExchange(ref _shutdownCompletionOwner, 2, 0) != 0)
            {
                return;
            }

            DisposeShutdownWatchdog();
            DisposeListener();
            RecordShutdownCompletion(_dispatcher?.QueueDepth ?? 0, Volatile.Read(ref _trackedTaskCount));
        }

        private void ArmShutdownWatchdog()
        {
            var watchdog = new Timer(
                _ => CompleteShutdownAfterOutlookThreadTimeout(),
                null,
                TimeSpan.FromSeconds(2),
                Timeout.InfiniteTimeSpan);
            var existing = Interlocked.CompareExchange(ref _shutdownWatchdog, watchdog, null);
            if (existing != null)
            {
                watchdog.Dispose();
                return;
            }

            if (Volatile.Read(ref _shutdownCompletionOwner) != 0)
            {
                DisposeShutdownWatchdog();
            }
        }

        private void CompleteShutdownAfterOutlookThreadTimeout()
        {
            if (Interlocked.CompareExchange(ref _shutdownCompletionOwner, 2, 0) != 0)
            {
                return;
            }

            RecordShutdownFailure(new InvalidOperationException("The Outlook STA did not complete final shutdown."));
            DisposeShutdownWatchdog();
            RecordShutdownCompletion(_dispatcher?.QueueDepth ?? 0, Volatile.Read(ref _trackedTaskCount));
        }

        private void DisposeShutdownWatchdog()
        {
            Interlocked.Exchange(ref _shutdownWatchdog, null)?.Dispose();
        }

        private void RecordShutdownCompletion(int queueDepth, int trackedTaskCount)
        {
            var failure = GetShutdownFailure();
            var durationTicks = Stopwatch.GetTimestamp() - _shutdownStartedTicks;
            if (failure == null &&
                _lifecycle.State == HostLifecycleState.Stopped &&
                queueDepth == 0 &&
                trackedTaskCount == 0)
            {
                _diagnostics.RecordShutdown(
                    _lifecycle.State,
                    durationTicks,
                    queueDepth,
                    trackedTaskCount);
            }
            else
            {
                _diagnostics.RecordFailure(
                    RuntimeDiagnosticEvent.ShutdownCompleted,
                    _lifecycle.State,
                    durationTicks,
                    queueDepth,
                    trackedTaskCount,
                    failure ?? new InvalidOperationException("HOST_NOT_QUIESCENT"));
            }

            _diagnostics.Dispose();
            _shutdownCompletion.TrySetResult(true);
        }

        private void ObserveInitialization(Task? initializationTask)
        {
            if (initializationTask == null)
            {
                return;
            }

            try
            {
                initializationTask.GetAwaiter().GetResult();
            }
            catch (Exception exception)
            {
                RecordShutdownFailure(exception);
            }
        }

        private OutlookStatusSnapshot CreateStatusSnapshot()
        {
            var version = typeof(AddInHost).Assembly.GetName().Version?.ToString() ?? "1.0.0.0";
            return new OutlookStatusSnapshot(
                _lifecycle.State.ToString().ToLowerInvariant(),
                IsListenerActive(),
                version);
        }

        private bool IsListenerActive()
        {
            lock (_transportGate)
            {
                return _server?.IsListening == true;
            }
        }

        private void BeginListenerShutdown()
        {
            LoopbackHttpServer? server;
            lock (_transportGate)
            {
                server = _server;
            }

            try
            {
                _ = server?.BeginShutdown();
            }
            catch (Exception exception)
            {
                RecordShutdownFailure(exception);
            }
        }

        private void ObserveListenerCompletion(LoopbackHttpServer server)
        {
            var completion = server.Completion;
            completion.ConfigureAwait(false).GetAwaiter().OnCompleted(
                () => HandleListenerCompletion(server, completion));
        }

        private void HandleListenerCompletion(LoopbackHttpServer server, Task completion)
        {
            Exception failure;
            try
            {
                completion.GetAwaiter().GetResult();
                failure = new InvalidOperationException("The loopback listener stopped unexpectedly.");
            }
            catch (Exception exception)
            {
                failure = exception;
            }

            if (Volatile.Read(ref _stopRequested) != 0)
            {
                return;
            }

            lock (_transportGate)
            {
                if (!ReferenceEquals(_server, server))
                {
                    return;
                }
            }

            if (Volatile.Read(ref _stopRequested) != 0 || !_lifecycle.TryMarkDegraded())
            {
                return;
            }

            _diagnostics.RecordFailure(
                RuntimeDiagnosticEvent.ListenerBindingCompleted,
                HostLifecycleState.Degraded,
                0,
                _dispatcher?.QueueDepth ?? 0,
                Volatile.Read(ref _trackedTaskCount),
                failure,
                listenerActive: false);
        }

        private void ObserveTransport(Task completion)
        {
            try
            {
                completion.GetAwaiter().GetResult();
            }
            catch (Exception exception)
            {
                RecordShutdownFailure(exception);
            }
        }

        private void DisposeListener()
        {
            LoopbackHttpServer? server;
            lock (_transportGate)
            {
                server = _server;
                _server = null;
            }

            if (server == null)
            {
                return;
            }

            try
            {
                server.Dispose();
            }
            catch (Exception exception)
            {
                RecordShutdownFailure(exception);
            }
        }

        private void RecordShutdownFailure(Exception exception)
        {
            lock (_shutdownGate)
            {
                _shutdownFailure ??= exception;
            }
        }

        private Exception? GetShutdownFailure()
        {
            lock (_shutdownGate)
            {
                return _shutdownFailure;
            }
        }
    }
}
