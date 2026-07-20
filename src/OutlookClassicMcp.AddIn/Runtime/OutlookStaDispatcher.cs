using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookClassicMcp.AddIn.Runtime
{
    internal sealed class OutlookStaDispatcher : IDisposable
    {
        private const int DefaultCapacity = 16;
        private readonly object _gate = new object();
        private readonly Queue<Action> _lifecycleCallbacks = new Queue<Action>();
        private readonly Queue<IDispatchWorkItem> _queue = new Queue<IDispatchWorkItem>();
        private readonly int _capacity;
        private readonly DispatchControl _control;
        private readonly Action? _onUnavailable;
        private readonly object? _ownerContext;
        private readonly Func<IntPtr, int, IntPtr, IntPtr, bool> _postMessage;
        private TaskCompletionSource<bool>? _shutdownQuiescence;
        private IntPtr _controlHandle;
        private bool _accepting = true;
        private bool _active;
        private bool _drainScheduled;
        private bool _draining;
        private bool _disposed;
        private bool _unavailableNotified;
        private bool _wakeupAvailable = true;

        public OutlookStaDispatcher()
            : this(null, DefaultCapacity, NativeMethods.PostMessage, null)
        {
        }

        internal OutlookStaDispatcher(object ownerContext)
            : this(
                ownerContext ?? throw new ArgumentNullException(nameof(ownerContext)),
                DefaultCapacity,
                NativeMethods.PostMessage,
                null)
        {
        }

        internal OutlookStaDispatcher(object ownerContext, Action onUnavailable)
            : this(
                ownerContext ?? throw new ArgumentNullException(nameof(ownerContext)),
                DefaultCapacity,
                NativeMethods.PostMessage,
                onUnavailable ?? throw new ArgumentNullException(nameof(onUnavailable)))
        {
        }

        internal OutlookStaDispatcher(Func<IntPtr, int, IntPtr, IntPtr, bool> postMessage)
            : this(null, DefaultCapacity, postMessage, null)
        {
        }

        internal OutlookStaDispatcher(
            object ownerContext,
            Func<IntPtr, int, IntPtr, IntPtr, bool> postMessage)
            : this(
                ownerContext ?? throw new ArgumentNullException(nameof(ownerContext)),
                DefaultCapacity,
                postMessage,
                null)
        {
        }

        internal OutlookStaDispatcher(
            object ownerContext,
            Func<IntPtr, int, IntPtr, IntPtr, bool> postMessage,
            Action onUnavailable)
            : this(
                ownerContext ?? throw new ArgumentNullException(nameof(ownerContext)),
                DefaultCapacity,
                postMessage,
                onUnavailable ?? throw new ArgumentNullException(nameof(onUnavailable)))
        {
        }

        private OutlookStaDispatcher(
            object? ownerContext,
            int capacity,
            Func<IntPtr, int, IntPtr, IntPtr, bool> postMessage,
            Action? onUnavailable)
        {
            if (capacity <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(capacity));
            }

            _postMessage = postMessage ?? throw new ArgumentNullException(nameof(postMessage));
            _onUnavailable = onUnavailable;
            _ownerContext = ownerContext;
            OwnerThread = OutlookThreadContext.Capture();
            if (OwnerThread.ApartmentState != ApartmentState.STA)
            {
                throw new InvalidOperationException("The Outlook dispatcher must be created on an STA thread.");
            }

            _capacity = capacity;
            _control = new DispatchControl(OnDispatchMessage);
            _controlHandle = _control.Handle;
            if (_controlHandle == IntPtr.Zero)
            {
                throw new InvalidOperationException("The Outlook dispatcher control did not create a window handle.");
            }
        }

        public OutlookThreadContext OwnerThread { get; }

        public bool IsAccepting
        {
            get
            {
                lock (_gate)
                {
                    return _accepting && !_disposed;
                }
            }
        }

        public int QueueDepth
        {
            get
            {
                lock (_gate)
                {
                    return _queue.Count + (_active ? 1 : 0);
                }
            }
        }

        public Task<T> InvokeAsync<T>(Func<T> operation, CancellationToken cancellationToken)
        {
            if (operation == null)
            {
                throw new ArgumentNullException(nameof(operation));
            }

            if (cancellationToken.IsCancellationRequested)
            {
                return Task.FromCanceled<T>(cancellationToken);
            }

            DispatchWorkItem<T> workItem;
            var scheduled = false;
            lock (_gate)
            {
                if (_disposed)
                {
                    return Task.FromException<T>(new InvalidOperationException("HOST_STOPPING"));
                }

                if (!_wakeupAvailable)
                {
                    return Task.FromException<T>(new InvalidOperationException("HOST_UNAVAILABLE"));
                }

                if (!_accepting)
                {
                    return Task.FromException<T>(new InvalidOperationException("HOST_STOPPING"));
                }

                if (_queue.Count + (_active ? 1 : 0) >= _capacity)
                {
                    return Task.FromException<T>(new InvalidOperationException("HOST_BUSY"));
                }

                workItem = new DispatchWorkItem<T>(operation, cancellationToken);
                _queue.Enqueue(workItem);
                scheduled = TryScheduleDrainUnderLock();
            }

            if (!scheduled)
            {
                FailDrain(null);
            }

            return workItem.Task;
        }

        public Task<TResult> InvokeWithContextAsync<TContext, TResult>(
            Func<TContext, TResult> operation,
            CancellationToken cancellationToken)
            where TContext : class
        {
            if (operation == null)
            {
                throw new ArgumentNullException(nameof(operation));
            }

            if (operation.Target != null)
            {
                throw new ArgumentException(
                    "Context-aware Outlook operations must be static.",
                    nameof(operation));
            }

            if (cancellationToken.IsCancellationRequested)
            {
                return Task.FromCanceled<TResult>(cancellationToken);
            }

            if (!(_ownerContext is TContext))
            {
                return Task.FromException<TResult>(new InvalidOperationException("HOST_UNAVAILABLE"));
            }

            ContextDispatchWorkItem<TContext, TResult> workItem;
            var scheduled = false;
            lock (_gate)
            {
                if (_disposed)
                {
                    return Task.FromException<TResult>(new InvalidOperationException("HOST_STOPPING"));
                }

                if (!_wakeupAvailable)
                {
                    return Task.FromException<TResult>(new InvalidOperationException("HOST_UNAVAILABLE"));
                }

                if (!_accepting)
                {
                    return Task.FromException<TResult>(new InvalidOperationException("HOST_STOPPING"));
                }

                if (_queue.Count + (_active ? 1 : 0) >= _capacity)
                {
                    return Task.FromException<TResult>(new InvalidOperationException("HOST_BUSY"));
                }

                workItem = new ContextDispatchWorkItem<TContext, TResult>(
                    operation,
                    cancellationToken);
                _queue.Enqueue(workItem);
                scheduled = TryScheduleDrainUnderLock();
            }

            if (!scheduled)
            {
                FailDrain(null);
            }

            return workItem.Task;
        }

        public Task<TResult> InvokeWithContextAsync<TContext, TState, TResult>(
            TState state,
            Func<TContext, TState, CancellationToken, TResult> operation,
            CancellationToken cancellationToken)
            where TContext : class
        {
            if (operation == null)
            {
                throw new ArgumentNullException(nameof(operation));
            }

            if (operation.Target != null)
            {
                throw new ArgumentException(
                    "Context-aware Outlook operations must be static.",
                    nameof(operation));
            }

            if (cancellationToken.IsCancellationRequested)
            {
                return Task.FromCanceled<TResult>(cancellationToken);
            }

            if (!(_ownerContext is TContext))
            {
                return Task.FromException<TResult>(new InvalidOperationException("HOST_UNAVAILABLE"));
            }

            ContextStateDispatchWorkItem<TContext, TState, TResult> workItem;
            var scheduled = false;
            lock (_gate)
            {
                if (_disposed)
                {
                    return Task.FromException<TResult>(new InvalidOperationException("HOST_STOPPING"));
                }

                if (!_wakeupAvailable)
                {
                    return Task.FromException<TResult>(new InvalidOperationException("HOST_UNAVAILABLE"));
                }

                if (!_accepting)
                {
                    return Task.FromException<TResult>(new InvalidOperationException("HOST_STOPPING"));
                }

                if (_queue.Count + (_active ? 1 : 0) >= _capacity)
                {
                    return Task.FromException<TResult>(new InvalidOperationException("HOST_BUSY"));
                }

                workItem = new ContextStateDispatchWorkItem<TContext, TState, TResult>(
                    state,
                    operation,
                    cancellationToken);
                _queue.Enqueue(workItem);
                scheduled = TryScheduleDrainUnderLock();
            }

            if (!scheduled)
            {
                FailDrain(null);
            }

            return workItem.Task;
        }

        public void AssertOutlookThread()
        {
            var current = OutlookThreadContext.Capture();
            if (current.ManagedThreadId != OwnerThread.ManagedThreadId ||
                current.NativeThreadId != OwnerThread.NativeThreadId ||
                current.ApartmentState != ApartmentState.STA)
            {
                throw new InvalidOperationException("Outlook work executed outside the captured UI STA.");
            }
        }

        public Task BeginShutdown()
        {
            IDispatchWorkItem[] pending;
            TaskCompletionSource<bool> quiescence;
            var completeQuiescence = false;
            lock (_gate)
            {
                if (_shutdownQuiescence == null)
                {
                    _shutdownQuiescence = new TaskCompletionSource<bool>(
                        TaskCreationOptions.RunContinuationsAsynchronously);
                }
                quiescence = _shutdownQuiescence;

                if (_accepting)
                {
                    _accepting = false;
                    pending = _queue.ToArray();
                    _queue.Clear();
                }
                else
                {
                    pending = Array.Empty<IDispatchWorkItem>();
                }

                completeQuiescence = !_active && _queue.Count == 0;
            }

            var exception = new InvalidOperationException("HOST_STOPPING");
            foreach (var workItem in pending)
            {
                workItem.Fail(exception);
            }

            if (completeQuiescence)
            {
                quiescence.TrySetResult(true);
            }

            return quiescence.Task;
        }

        public bool TryPostLifecycleCallback(Action callback)
        {
            if (callback == null)
            {
                throw new ArgumentNullException(nameof(callback));
            }

            var scheduled = false;
            lock (_gate)
            {
                if (_disposed || !_wakeupAvailable)
                {
                    return false;
                }

                _lifecycleCallbacks.Enqueue(callback);
                scheduled = TryScheduleDrainUnderLock();
            }

            if (!scheduled)
            {
                FailDrain(null);
            }

            return scheduled;
        }

        public void Dispose()
        {
            AssertOutlookThread();
            var quiescence = BeginShutdown();
            if (!quiescence.IsCompleted)
            {
                throw new InvalidOperationException("HOST_NOT_QUIESCENT");
            }

            quiescence.GetAwaiter().GetResult();
            IntPtr controlHandle;
            lock (_gate)
            {
                if (_disposed)
                {
                    return;
                }
            }

            lock (_gate)
            {
                if (_disposed)
                {
                    return;
                }

                _disposed = true;
                _drainScheduled = false;
                _lifecycleCallbacks.Clear();
                _wakeupAvailable = false;
                controlHandle = _controlHandle;
                _controlHandle = IntPtr.Zero;
            }

            while (NativeMethods.PeekMessage(
                out _,
                controlHandle,
                DispatchControl.DispatchMessage,
                DispatchControl.DispatchMessage,
                NativeMethods.RemoveMessage))
            {
            }

            _control.Dispose();
        }

        private void OnDispatchMessage()
        {
            Action? lifecycleCallback = null;
            IDispatchWorkItem? workItem = null;
            var ownsDrain = false;
            try
            {
                lock (_gate)
                {
                    _drainScheduled = false;
                    if (_disposed || _draining)
                    {
                        return;
                    }

                    _draining = true;
                    ownsDrain = true;
                    if (_lifecycleCallbacks.Count > 0)
                    {
                        lifecycleCallback = _lifecycleCallbacks.Dequeue();
                    }
                    else if (_accepting && _queue.Count > 0)
                    {
                        workItem = _queue.Dequeue();
                        _active = true;
                    }
                }

                if (lifecycleCallback != null)
                {
                    AssertOutlookThread();
                    lifecycleCallback();
                    return;
                }

                if (workItem == null)
                {
                    return;
                }

                AssertOutlookThread();
                workItem.Invoke(_ownerContext);

                TaskCompletionSource<bool>? quiescence = null;
                lock (_gate)
                {
                    _active = false;
                    if (!_accepting && _queue.Count == 0)
                    {
                        quiescence = _shutdownQuiescence;
                    }
                }

                workItem.Complete();
                quiescence?.TrySetResult(true);
            }
            catch
            {
                FailDrain(workItem);
            }
            finally
            {
                if (ownsDrain)
                {
                    var scheduled = true;
                    lock (_gate)
                    {
                        _draining = false;
                        if (!_disposed &&
                            (_lifecycleCallbacks.Count > 0 || (_accepting && _queue.Count > 0)))
                        {
                            scheduled = TryScheduleDrainUnderLock();
                        }
                    }

                    if (!scheduled)
                    {
                        FailDrain(null);
                    }
                }
            }
        }

        private bool TryScheduleDrainUnderLock()
        {
            if (_disposed || !_wakeupAvailable)
            {
                return false;
            }

            if (_controlHandle == IntPtr.Zero)
            {
                _accepting = false;
                _lifecycleCallbacks.Clear();
                _wakeupAvailable = false;
                return false;
            }

            if (_drainScheduled)
            {
                return true;
            }

            _drainScheduled = true;
            if (_postMessage(
                _controlHandle,
                DispatchControl.DispatchMessage,
                IntPtr.Zero,
                IntPtr.Zero))
            {
                return true;
            }

            _drainScheduled = false;
            _accepting = false;
            _lifecycleCallbacks.Clear();
            _wakeupAvailable = false;
            return false;
        }

        private void FailDrain(IDispatchWorkItem? activeWorkItem)
        {
            IDispatchWorkItem[] pending;
            TaskCompletionSource<bool>? quiescence;
            var notifyUnavailable = false;
            lock (_gate)
            {
                _accepting = false;
                if (activeWorkItem != null)
                {
                    _active = false;
                }
                _drainScheduled = false;
                _wakeupAvailable = false;
                _lifecycleCallbacks.Clear();
                pending = _queue.ToArray();
                _queue.Clear();
                quiescence = !_active && _queue.Count == 0
                    ? _shutdownQuiescence
                    : null;
                if (!_disposed && !_unavailableNotified)
                {
                    _unavailableNotified = true;
                    notifyUnavailable = true;
                }
            }

            var exception = new InvalidOperationException("HOST_UNAVAILABLE");
            activeWorkItem?.Fail(exception);
            foreach (var workItem in pending)
            {
                workItem.Fail(exception);
            }

            quiescence?.TrySetResult(true);
            if (notifyUnavailable)
            {
                try
                {
                    _onUnavailable?.Invoke();
                }
                catch (Exception)
                {
                }
            }
        }

        private sealed class DispatchControl : System.Windows.Forms.Control
        {
            internal const int DispatchMessage = 0x8001;
            private readonly Action _dispatch;

            public DispatchControl(Action dispatch)
            {
                _dispatch = dispatch;
            }

            protected override void WndProc(ref System.Windows.Forms.Message message)
            {
                if (message.Msg == DispatchMessage)
                {
                    try
                    {
                        _dispatch();
                    }
                    catch
                    {
                    }

                    return;
                }

                base.WndProc(ref message);
            }
        }

        private static class NativeMethods
        {
            internal const uint RemoveMessage = 0x0001;

            [DllImport("user32.dll", SetLastError = true)]
            [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool PostMessage(
                IntPtr windowHandle,
                int message,
                IntPtr wordParameter,
                IntPtr longParameter);

            [DllImport("user32.dll", SetLastError = true)]
            [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool PeekMessage(
                out NativeMessage message,
                IntPtr windowHandle,
                uint minimumMessage,
                uint maximumMessage,
                uint removeMessage);

            [StructLayout(LayoutKind.Sequential)]
            internal struct NativeMessage
            {
                internal IntPtr WindowHandle;
                internal uint Message;
                internal UIntPtr WordParameter;
                internal IntPtr LongParameter;
                internal uint Time;
                internal NativePoint Point;
            }

            [StructLayout(LayoutKind.Sequential)]
            internal struct NativePoint
            {
                internal int X;
                internal int Y;
            }
        }

        private interface IDispatchWorkItem
        {
            void Invoke(object? ownerContext);

            void Complete();

            void Fail(Exception exception);
        }

        private sealed class DispatchWorkItem<T> : IDispatchWorkItem
        {
            private const int Queued = 0;
            private const int Canceled = 1;
            private const int Running = 2;
            private const int Invoked = 3;
            private const int Terminal = 4;
            private readonly CancellationToken _cancellationToken;
            private readonly CancellationTokenRegistration _cancellationRegistration;
            private readonly Func<T> _operation;
            private readonly TaskCompletionSource<T> _completion =
                new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
            private Exception? _exception;
            private T _result = default!;
            private int _state = Queued;

            public DispatchWorkItem(Func<T> operation, CancellationToken cancellationToken)
            {
                _operation = operation;
                _cancellationToken = cancellationToken;
                _cancellationRegistration = cancellationToken.Register(CancelQueuedWork);
            }

            public Task<T> Task => _completion.Task;

            public void Invoke(object? ownerContext)
            {
                if (Interlocked.CompareExchange(ref _state, Running, Queued) != Queued)
                {
                    return;
                }

                try
                {
                    _result = _operation();
                }
                catch (Exception exception)
                {
                    _exception = exception;
                }
                finally
                {
                    Volatile.Write(ref _state, Invoked);
                }
            }

            public void Complete()
            {
                var state = Volatile.Read(ref _state);
                if (state == Canceled)
                {
                    Interlocked.CompareExchange(ref _state, Terminal, Canceled);
                }
                else if (Interlocked.CompareExchange(ref _state, Terminal, Invoked) == Invoked)
                {
                    if (_exception != null)
                    {
                        _completion.TrySetException(_exception);
                    }
                    else
                    {
                        _completion.TrySetResult(_result);
                    }
                }

                _cancellationRegistration.Dispose();
            }

            public void Fail(Exception exception)
            {
                while (true)
                {
                    var state = Volatile.Read(ref _state);
                    if (state == Terminal)
                    {
                        break;
                    }

                    if (Interlocked.CompareExchange(ref _state, Terminal, state) != state)
                    {
                        continue;
                    }

                    if (state != Canceled)
                    {
                        _completion.TrySetException(exception);
                    }

                    break;
                }

                _cancellationRegistration.Dispose();
            }

            private void CancelQueuedWork()
            {
                if (Interlocked.CompareExchange(ref _state, Canceled, Queued) == Queued)
                {
                    _completion.TrySetCanceled(_cancellationToken);
                }
            }
        }

        private sealed class ContextDispatchWorkItem<TContext, TResult> : IDispatchWorkItem
            where TContext : class
        {
            private const int Queued = 0;
            private const int Canceled = 1;
            private const int Running = 2;
            private const int Invoked = 3;
            private const int Terminal = 4;
            private readonly CancellationToken _cancellationToken;
            private readonly CancellationTokenRegistration _cancellationRegistration;
            private readonly TaskCompletionSource<TResult> _completion =
                new TaskCompletionSource<TResult>(TaskCreationOptions.RunContinuationsAsynchronously);
            private readonly Func<TContext, TResult> _operation;
            private Exception? _exception;
            private TResult _result = default!;
            private int _state = Queued;

            public ContextDispatchWorkItem(
                Func<TContext, TResult> operation,
                CancellationToken cancellationToken)
            {
                _operation = operation;
                _cancellationToken = cancellationToken;
                _cancellationRegistration = cancellationToken.Register(CancelQueuedWork);
            }

            public Task<TResult> Task => _completion.Task;

            public void Invoke(object? ownerContext)
            {
                if (Interlocked.CompareExchange(ref _state, Running, Queued) != Queued)
                {
                    return;
                }

                try
                {
                    if (!(ownerContext is TContext typedContext))
                    {
                        throw new InvalidOperationException("HOST_UNAVAILABLE");
                    }

                    _result = _operation(typedContext);
                }
                catch (Exception exception)
                {
                    _exception = exception;
                }
                finally
                {
                    Volatile.Write(ref _state, Invoked);
                }
            }

            public void Complete()
            {
                var state = Volatile.Read(ref _state);
                if (state == Canceled)
                {
                    Interlocked.CompareExchange(ref _state, Terminal, Canceled);
                }
                else if (Interlocked.CompareExchange(ref _state, Terminal, Invoked) == Invoked)
                {
                    if (_exception != null)
                    {
                        _completion.TrySetException(_exception);
                    }
                    else
                    {
                        _completion.TrySetResult(_result);
                    }
                }

                _cancellationRegistration.Dispose();
            }

            public void Fail(Exception exception)
            {
                while (true)
                {
                    var state = Volatile.Read(ref _state);
                    if (state == Terminal)
                    {
                        break;
                    }

                    if (Interlocked.CompareExchange(ref _state, Terminal, state) != state)
                    {
                        continue;
                    }

                    if (state != Canceled)
                    {
                        _completion.TrySetException(exception);
                    }

                    break;
                }

                _cancellationRegistration.Dispose();
            }

            private void CancelQueuedWork()
            {
                if (Interlocked.CompareExchange(ref _state, Canceled, Queued) == Queued)
                {
                    _completion.TrySetCanceled(_cancellationToken);
                }
            }
        }

        private sealed class ContextStateDispatchWorkItem<TContext, TState, TResult> : IDispatchWorkItem
            where TContext : class
        {
            private const int Queued = 0;
            private const int Canceled = 1;
            private const int Running = 2;
            private const int Invoked = 3;
            private const int Terminal = 4;
            private readonly CancellationToken _cancellationToken;
            private readonly CancellationTokenRegistration _cancellationRegistration;
            private readonly TaskCompletionSource<TResult> _completion =
                new TaskCompletionSource<TResult>(TaskCreationOptions.RunContinuationsAsynchronously);
            private readonly Func<TContext, TState, CancellationToken, TResult> _operation;
            private readonly TState _operationState;
            private Exception? _exception;
            private TResult _result = default!;
            private int _state = Queued;

            public ContextStateDispatchWorkItem(
                TState operationState,
                Func<TContext, TState, CancellationToken, TResult> operation,
                CancellationToken cancellationToken)
            {
                _operationState = operationState;
                _operation = operation;
                _cancellationToken = cancellationToken;
                _cancellationRegistration = cancellationToken.Register(CancelQueuedWork);
            }

            public Task<TResult> Task => _completion.Task;

            public void Invoke(object? ownerContext)
            {
                if (Interlocked.CompareExchange(ref _state, Running, Queued) != Queued)
                {
                    return;
                }

                try
                {
                    if (!(ownerContext is TContext typedContext))
                    {
                        throw new InvalidOperationException("HOST_UNAVAILABLE");
                    }

                    _cancellationToken.ThrowIfCancellationRequested();
                    _result = _operation(typedContext, _operationState, _cancellationToken);
                }
                catch (Exception exception)
                {
                    _exception = exception;
                }
                finally
                {
                    Volatile.Write(ref _state, Invoked);
                }
            }

            public void Complete()
            {
                var state = Volatile.Read(ref _state);
                if (state == Canceled)
                {
                    Interlocked.CompareExchange(ref _state, Terminal, Canceled);
                }
                else if (Interlocked.CompareExchange(ref _state, Terminal, Invoked) == Invoked)
                {
                    if (_exception != null)
                    {
                        _completion.TrySetException(_exception);
                    }
                    else
                    {
                        _completion.TrySetResult(_result);
                    }
                }

                _cancellationRegistration.Dispose();
            }

            public void Fail(Exception exception)
            {
                while (true)
                {
                    var state = Volatile.Read(ref _state);
                    if (state == Terminal)
                    {
                        break;
                    }

                    if (Interlocked.CompareExchange(ref _state, Terminal, state) != state)
                    {
                        continue;
                    }

                    if (state != Canceled)
                    {
                        _completion.TrySetException(exception);
                    }

                    break;
                }

                _cancellationRegistration.Dispose();
            }

            private void CancelQueuedWork()
            {
                if (Interlocked.CompareExchange(ref _state, Canceled, Queued) == Queued)
                {
                    _completion.TrySetCanceled(_cancellationToken);
                }
            }
        }
    }
}
