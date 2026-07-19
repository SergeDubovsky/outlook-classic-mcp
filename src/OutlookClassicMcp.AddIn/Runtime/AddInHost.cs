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
        private readonly Outlook.ApplicationEvents_11_Event _applicationEvents;
        private readonly MetadataDiagnostics _diagnostics;
        private readonly HostLifecycle _lifecycle = new HostLifecycle();
        private readonly object _shutdownGate = new object();
        private readonly TaskCompletionSource<bool> _shutdownCompletion =
            new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        private OutlookStaDispatcher? _dispatcher;
        private Task? _initializationTask;
        private Timer? _shutdownWatchdog;
        private Exception? _shutdownFailure;
        private Exception? _startupFailure;
        private long _shutdownStartedTicks;
        private int _shutdownCompletionOwner;
        private int _staCleanupStarted;
        private int _trackedTaskCount;
        private bool _quitSubscribed;

        public AddInHost(Outlook.Application application)
        {
            _ = application ?? throw new ArgumentNullException(nameof(application));
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
                if (dispatcher == null || !dispatcher.IsAccepting)
                {
                    _dispatcher = null;
                    dispatcher?.Dispose();
                    dispatcher = new OutlookStaDispatcher();
                    _dispatcher = dispatcher;
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

            var initializationTask = _initializationTask;
            if (initializationTask == null || initializationTask.IsCompleted)
            {
                ObserveInitialization(initializationTask);
                FinalizeShutdownOnOutlookThread();
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
                    trackedTaskCount);
            }
            else
            {
                _diagnostics.RecordFailure(
                    RuntimeDiagnosticEvent.StartupCompleted,
                    _lifecycle.State,
                    durationTicks,
                    queueDepth,
                    trackedTaskCount,
                    failure);
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
                _lifecycle.TryMarkOnline();
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
                    exception);
            }
            finally
            {
                var remainingTasks = Interlocked.Decrement(ref _trackedTaskCount);
                _diagnostics.RecordHostQuiescent(
                    _lifecycle.State,
                    _dispatcher?.QueueDepth ?? 0,
                    remainingTasks);
            }
        }

        private void OnApplicationQuit()
        {
            _ = BeginShutdown();
        }

        private void CompleteShutdownAfterInitialization(Task initializationTask)
        {
            ObserveInitialization(initializationTask);
            var dispatcher = _dispatcher;
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
