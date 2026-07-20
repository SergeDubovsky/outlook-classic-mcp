#if NET48
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using NUnit.Framework;
using OutlookClassicMcp.AddIn.Runtime;

namespace OutlookClassicMcp.Core.Tests
{
    [TestFixture]
    public sealed class OutlookStaDispatcherContextTests
    {
        private const int DefaultCapacity = 16;
        private static readonly int[] ExpectedFirstExecution = { 1 };
        private static readonly int[] ExpectedFirstTwoExecutions = { 1, 2 };

        [Test]
        public void StatefulContextWorkReceivesManagedStateAndCancellationTokenOnOwnerSta()
        {
            var context = new DispatcherOwnerContext();
            using (var harness = new StaDispatcherHarness(context))
            {
                var result = harness.Dispatcher.InvokeWithContextAsync<
                    DispatcherOwnerContext,
                    string,
                    string>(
                    "managed-state",
                    ExecuteStateful,
                    CancellationToken.None);

                Assert.That(result.Wait(TimeSpan.FromSeconds(5)), Is.True);
                Assert.That(result.Result, Is.EqualTo("managed-state"));
                Assert.That(context.ExecutedThreads, Has.Count.EqualTo(1));
                Assert.That(
                    context.ExecutedThreads[0].ManagedThreadId,
                    Is.EqualTo(harness.Dispatcher.OwnerThread.ManagedThreadId));
                Assert.That(context.ExecutedThreads[0].ApartmentState, Is.EqualTo(ApartmentState.STA));
            }
        }

        [Test]
        public void StatefulContextInvocationRejectsBoundDelegateBeforeQueueing()
        {
            var context = new DispatcherOwnerContext();
            using (var harness = new StaDispatcherHarness(context))
            {
                Func<DispatcherOwnerContext, string, CancellationToken, string> boundOperation =
                    context.BoundStateOperation;
                Action invokeBoundOperation = () =>
                    harness.Dispatcher.InvokeWithContextAsync(
                        "managed-state",
                        boundOperation,
                        CancellationToken.None);

                Assert.Throws<ArgumentException>(invokeBoundOperation);
                Assert.That(harness.Dispatcher.QueueDepth, Is.Zero);
            }
        }

        [Test]
        public void PreCanceledStatefulContextWorkNeverQueues()
        {
            var context = new DispatcherOwnerContext();
            using (var cancellation = new CancellationTokenSource())
            using (var harness = new StaDispatcherHarness(context))
            {
                cancellation.Cancel();
                var result = harness.Dispatcher.InvokeWithContextAsync<
                    DispatcherOwnerContext,
                    string,
                    string>(
                    "managed-state",
                    ExecuteStateful,
                    cancellation.Token);

                Assert.That(result.IsCanceled, Is.True);
                Assert.That(harness.Dispatcher.QueueDepth, Is.Zero);
                Assert.That(context.ExecutedThreads, Is.Empty);
            }
        }

        [Test]
        public void CancellationWhileStatefulContextWorkIsQueuedPreventsExecution()
        {
            var context = new DispatcherOwnerContext();
            using (var cancellation = new CancellationTokenSource())
            using (var harness = new StaDispatcherHarness(context))
            {
                var active = harness.Dispatcher.InvokeWithContextAsync<
                    DispatcherOwnerContext,
                    int,
                    int>(1, ExecuteBlockingStateful, CancellationToken.None);
                Assert.That(context.FirstEntered.Wait(TimeSpan.FromSeconds(5)), Is.True);

                var queued = harness.Dispatcher.InvokeWithContextAsync<
                    DispatcherOwnerContext,
                    int,
                    int>(2, ExecuteStatefulInt, cancellation.Token);
                cancellation.Cancel();

                Assert.That(
                    SpinWait.SpinUntil(() => queued.IsCompleted, TimeSpan.FromSeconds(5)),
                    Is.True);
                Assert.That(queued.IsCanceled, Is.True);
                context.ReleaseFirst.Set();
                Assert.That(active.Wait(TimeSpan.FromSeconds(5)), Is.True);
                Assert.That(active.Result, Is.EqualTo(1));
                Assert.That(context.ExecutionCount, Is.EqualTo(1));
            }
        }

        [Test]
        public void RunningStatefulContextWorkObservesCooperativeCancellation()
        {
            var context = new DispatcherOwnerContext();
            using (var cancellation = new CancellationTokenSource())
            using (var harness = new StaDispatcherHarness(context))
            {
                var active = harness.Dispatcher.InvokeWithContextAsync<
                    DispatcherOwnerContext,
                    int,
                    int>(1, ExecuteCooperativelyCanceled, cancellation.Token);
                Assert.That(context.FirstEntered.Wait(TimeSpan.FromSeconds(5)), Is.True);

                cancellation.Cancel();

                Assert.That(
                    SpinWait.SpinUntil(() => active.IsCompleted, TimeSpan.FromSeconds(5)),
                    Is.True);
                Assert.That(active.IsFaulted, Is.True);
                Assert.That(
                    active.Exception!.Flatten().InnerException,
                    Is.TypeOf<OperationCanceledException>());
            }
        }

        [Test]
        public void StatefulContextShutdownWaitsForActiveAndFailsQueuedWork()
        {
            var context = new DispatcherOwnerContext();
            using (var harness = new StaDispatcherHarness(context))
            {
                var active = harness.Dispatcher.InvokeWithContextAsync<
                    DispatcherOwnerContext,
                    int,
                    int>(1, ExecuteBlockingStateful, CancellationToken.None);
                Assert.That(context.FirstEntered.Wait(TimeSpan.FromSeconds(5)), Is.True);
                var queued = harness.Dispatcher.InvokeWithContextAsync<
                    DispatcherOwnerContext,
                    int,
                    int>(2, ExecuteStatefulInt, CancellationToken.None);

                var shutdown = harness.Dispatcher.BeginShutdown();

                AssertFaultedWithMessage(queued, "HOST_STOPPING");
                Assert.That(shutdown.IsCompleted, Is.False);
                context.ReleaseFirst.Set();
                Assert.That(active.Wait(TimeSpan.FromSeconds(5)), Is.True);
                Assert.That(shutdown.Wait(TimeSpan.FromSeconds(5)), Is.True);
                Assert.That(active.Result, Is.EqualTo(1));
                Assert.That(context.ExecutionCount, Is.EqualTo(1));
            }
        }

        [Test]
        public void StatefulContextOperationExceptionCompletesTheTaskAsFaulted()
        {
            var context = new DispatcherOwnerContext();
            using (var harness = new StaDispatcherHarness(context))
            {
                var result = harness.Dispatcher.InvokeWithContextAsync<
                    DispatcherOwnerContext,
                    int,
                    int>(1, ThrowStateful, CancellationToken.None);

                Assert.That(
                    SpinWait.SpinUntil(() => result.IsCompleted, TimeSpan.FromSeconds(5)),
                    Is.True);
                Assert.That(result.IsFaulted, Is.True);
                Assert.That(
                    result.Exception!.Flatten().InnerException,
                    Is.TypeOf<InvalidOperationException>());
            }
        }

        [Test]
        public void ContextWorkIsSerializedOnOwnerStaAndQueuedStateIsIsolated()
        {
            var context = new DispatcherOwnerContext();
            using (var harness = new StaDispatcherHarness(context))
            {
                var first = harness.Dispatcher.InvokeWithContextAsync<DispatcherOwnerContext, int>(
                    ExecuteBlockingFirst,
                    CancellationToken.None);
                Assert.That(context.FirstEntered.Wait(TimeSpan.FromSeconds(5)), Is.True);

                var second = harness.Dispatcher.InvokeWithContextAsync<DispatcherOwnerContext, int>(
                    ExecuteBlockingFirst,
                    CancellationToken.None);
                Thread.Sleep(100);
                Assert.That(second.IsCompleted, Is.False);

                var queuedItem = GetQueuedItems(harness.Dispatcher).Single();
                AssertQueuedContextWorkItemIsIsolated(queuedItem, context);

                context.ReleaseFirst.Set();
                Assert.That(Task.WaitAll(new Task[] { first, second }, TimeSpan.FromSeconds(5)), Is.True);
                Assert.That(first.Result, Is.EqualTo(1));
                Assert.That(second.Result, Is.EqualTo(2));
                Assert.That(context.ExecutionOrder, Is.EqualTo(ExpectedFirstTwoExecutions));
                Assert.That(context.MaximumConcurrentExecutions, Is.EqualTo(1));
                Assert.That(context.ExecutedThreads, Has.Count.EqualTo(2));
                Assert.That(
                    context.ExecutedThreads.All(
                        thread => thread.ManagedThreadId == harness.Dispatcher.OwnerThread.ManagedThreadId &&
                            thread.NativeThreadId == harness.Dispatcher.OwnerThread.NativeThreadId &&
                            thread.ApartmentState == ApartmentState.STA),
                    Is.True);
            }
        }

        [Test]
        public void ContextInvocationRejectsBoundDelegateBeforeQueueing()
        {
            var context = new DispatcherOwnerContext();
            using (var harness = new StaDispatcherHarness(context))
            {
                Func<DispatcherOwnerContext, int> boundOperation = context.BoundOperation;
                Action invokeBoundOperation = () =>
                    harness.Dispatcher.InvokeWithContextAsync(
                        boundOperation,
                        CancellationToken.None);
                Assert.Throws<ArgumentException>(invokeBoundOperation);
                Assert.That(harness.Dispatcher.QueueDepth, Is.Zero);
            }
        }

        [Test]
        public void CancellationWhileQueuedPreventsContextOperation()
        {
            var context = new DispatcherOwnerContext();
            using (var cancellation = new CancellationTokenSource())
            using (var harness = new StaDispatcherHarness(context))
            {
                var active = harness.Dispatcher.InvokeWithContextAsync<DispatcherOwnerContext, int>(
                    ExecuteBlockingFirst,
                    CancellationToken.None);
                Assert.That(context.FirstEntered.Wait(TimeSpan.FromSeconds(5)), Is.True);

                var queued = harness.Dispatcher.InvokeWithContextAsync<DispatcherOwnerContext, int>(
                    ExecuteBlockingFirst,
                    cancellation.Token);
                cancellation.Cancel();
                Assert.That(
                    SpinWait.SpinUntil(() => queued.IsCompleted, TimeSpan.FromSeconds(5)),
                    Is.True);
                Assert.That(queued.IsCanceled, Is.True);
                Assert.That(context.ExecutionCount, Is.EqualTo(1));

                context.ReleaseFirst.Set();
                Assert.That(active.Wait(TimeSpan.FromSeconds(5)), Is.True);
                var recovery = harness.Dispatcher.InvokeWithContextAsync<DispatcherOwnerContext, int>(
                    ExecuteBlockingFirst,
                    CancellationToken.None);
                Assert.That(recovery.Wait(TimeSpan.FromSeconds(5)), Is.True);
                Assert.That(recovery.Result, Is.EqualTo(2));
                Assert.That(context.ExecutionCount, Is.EqualTo(2));
            }
        }

        [Test]
        public void CancellationAfterContextOperationStartsDoesNotForceAbortIt()
        {
            var context = new DispatcherOwnerContext();
            using (var cancellation = new CancellationTokenSource())
            using (var harness = new StaDispatcherHarness(context))
            {
                var active = harness.Dispatcher.InvokeWithContextAsync<DispatcherOwnerContext, int>(
                    ExecuteBlockingFirst,
                    cancellation.Token);
                Assert.That(context.FirstEntered.Wait(TimeSpan.FromSeconds(5)), Is.True);

                cancellation.Cancel();
                Thread.Sleep(100);
                Assert.That(active.IsCompleted, Is.False);

                context.ReleaseFirst.Set();
                Assert.That(active.Wait(TimeSpan.FromSeconds(5)), Is.True);
                Assert.That(active.Status, Is.EqualTo(TaskStatus.RanToCompletion));
                Assert.That(active.Result, Is.EqualTo(1));
            }
        }

        [Test]
        public void ContextShutdownWaitsForActiveOperationAndTerminatesQueuedWork()
        {
            var context = new DispatcherOwnerContext();
            using (var activeCancellation = new CancellationTokenSource())
            using (var queuedCancellation = new CancellationTokenSource())
            using (var harness = new StaDispatcherHarness(context))
            {
                var active = harness.Dispatcher.InvokeWithContextAsync<DispatcherOwnerContext, int>(
                    ExecuteBlockingFirst,
                    activeCancellation.Token);
                Assert.That(context.FirstEntered.Wait(TimeSpan.FromSeconds(5)), Is.True);

                activeCancellation.Cancel();
                Thread.Sleep(100);
                Assert.That(active.IsCompleted, Is.False);

                var queued = harness.Dispatcher.InvokeWithContextAsync<DispatcherOwnerContext, int>(
                    ExecuteBlockingFirst,
                    CancellationToken.None);
                var canceled = harness.Dispatcher.InvokeWithContextAsync<DispatcherOwnerContext, int>(
                    ExecuteBlockingFirst,
                    queuedCancellation.Token);
                queuedCancellation.Cancel();
                Assert.That(
                    SpinWait.SpinUntil(() => canceled.IsCompleted, TimeSpan.FromSeconds(5)),
                    Is.True);
                Assert.That(canceled.IsCanceled, Is.True);

                var firstShutdown = harness.Dispatcher.BeginShutdown();
                var repeatedShutdown = harness.Dispatcher.BeginShutdown();

                Assert.That(repeatedShutdown, Is.SameAs(firstShutdown));
                Assert.That(firstShutdown.IsCompleted, Is.False);
                AssertFaultedWithMessage(queued, "HOST_STOPPING");
                Assert.That(canceled.IsCanceled, Is.True);
                Assert.That(active.IsCompleted, Is.False);

                context.ReleaseFirst.Set();
                Assert.That(active.Wait(TimeSpan.FromSeconds(5)), Is.True);
                Assert.That(firstShutdown.Wait(TimeSpan.FromSeconds(5)), Is.True);
                Assert.That(active.Status, Is.EqualTo(TaskStatus.RanToCompletion));
                Assert.That(active.Result, Is.EqualTo(1));
                Assert.That(context.ExecutionCount, Is.EqualTo(1));
                Assert.That(context.ExecutionOrder, Is.EqualTo(ExpectedFirstExecution));
                Assert.That(context.ExecutedThreads, Has.Count.EqualTo(1));
                Assert.That(harness.Dispatcher.QueueDepth, Is.Zero);
                Assert.That(harness.Dispatcher.BeginShutdown(), Is.SameAs(firstShutdown));
            }
        }

        [Test]
        public void ContextQueueCapacityRecoversAfterAcceptedWorkDrains()
        {
            var context = new DispatcherOwnerContext();
            using (var harness = new StaDispatcherHarness(context))
            {
                var accepted = new List<Task<int>>(DefaultCapacity)
                {
                    harness.Dispatcher.InvokeWithContextAsync<DispatcherOwnerContext, int>(
                        ExecuteBlockingFirst,
                        CancellationToken.None),
                };
                Assert.That(context.FirstEntered.Wait(TimeSpan.FromSeconds(5)), Is.True);

                for (var index = 1; index < DefaultCapacity; index++)
                {
                    accepted.Add(
                        harness.Dispatcher.InvokeWithContextAsync<DispatcherOwnerContext, int>(
                            ExecuteBlockingFirst,
                            CancellationToken.None));
                }

                Assert.That(harness.Dispatcher.QueueDepth, Is.EqualTo(DefaultCapacity));
                var overflow = harness.Dispatcher.InvokeWithContextAsync<
                    DispatcherOwnerContext,
                    int,
                    int>(
                    99,
                    ExecuteStatefulInt,
                    CancellationToken.None);
                AssertFaultedWithMessage(overflow, "HOST_BUSY");

                context.ReleaseFirst.Set();
                Assert.That(Task.WaitAll(accepted.Cast<Task>().ToArray(), TimeSpan.FromSeconds(5)), Is.True);
                Assert.That(context.ExecutionCount, Is.EqualTo(DefaultCapacity));
                Assert.That(context.MaximumConcurrentExecutions, Is.EqualTo(1));

                var recovered = harness.Dispatcher.InvokeWithContextAsync<DispatcherOwnerContext, int>(
                    ExecuteBlockingFirst,
                    CancellationToken.None);
                Assert.That(recovered.Wait(TimeSpan.FromSeconds(5)), Is.True);
                Assert.That(recovered.Result, Is.EqualTo(DefaultCapacity + 1));
            }
        }

        [Test]
        public void ContextStoppingAndWakeupFailureRemainDistinct()
        {
            var stoppingContext = new DispatcherOwnerContext();
            using (var harness = new StaDispatcherHarness(stoppingContext))
            {
                harness.Dispatcher.BeginShutdown();
                var stopped = harness.Dispatcher.InvokeWithContextAsync<
                    DispatcherOwnerContext,
                    int,
                    int>(
                    1,
                    ExecuteStatefulInt,
                    CancellationToken.None);
                AssertFaultedWithMessage(stopped, "HOST_STOPPING");
            }

            Task<int>? unavailable = null;
            Task<int>? unavailableAgain = null;
            Task? unavailableShutdown = null;
            Exception? uiThreadFailure = null;
            var unavailableNotifications = 0;
            using (var unavailableContext = new DispatcherOwnerContext())
            {
                var uiThread = new Thread(() =>
                {
                    try
                    {
                        using (var dispatcher = new OutlookStaDispatcher(
                            unavailableContext,
                            (_, _, _, _) => false,
                            () => Interlocked.Increment(ref unavailableNotifications)))
                        {
                            unavailable = dispatcher.InvokeWithContextAsync<DispatcherOwnerContext, int>(
                                ExecuteBlockingFirst,
                                CancellationToken.None);
                            unavailableAgain = dispatcher.InvokeWithContextAsync<
                                DispatcherOwnerContext,
                                int,
                                int>(
                                1,
                                ExecuteStatefulInt,
                                CancellationToken.None);
                            unavailableShutdown = dispatcher.BeginShutdown();
                        }
                    }
                    catch (Exception exception)
                    {
                        uiThreadFailure = exception;
                    }
                });
                uiThread.SetApartmentState(ApartmentState.STA);
                uiThread.Start();

                Assert.That(uiThread.Join(TimeSpan.FromSeconds(5)), Is.True);
                Assert.That(uiThreadFailure, Is.Null);
                Assert.That(unavailable, Is.Not.Null);
                AssertFaultedWithMessage(unavailable!, "HOST_UNAVAILABLE");
                Assert.That(unavailableAgain, Is.Not.Null);
                AssertFaultedWithMessage(unavailableAgain!, "HOST_UNAVAILABLE");
                Assert.That(unavailableShutdown, Is.Not.Null);
                Assert.That(unavailableShutdown!.Status, Is.EqualTo(TaskStatus.RanToCompletion));
                Assert.That(unavailableNotifications, Is.EqualTo(1));
            }
        }

        private static int ExecuteBlockingFirst(DispatcherOwnerContext context)
        {
            var execution = Interlocked.Increment(ref context.ExecutionCount);
            var active = Interlocked.Increment(ref context.ActiveExecutions);
            UpdateMaximum(ref context.MaximumConcurrentExecutions, active);
            try
            {
                if (execution == 1)
                {
                    context.FirstEntered.Set();
                    while (!context.ReleaseFirst.IsSet)
                    {
                        Application.DoEvents();
                        Thread.Sleep(1);
                    }
                }

                context.ExecutedThreads.Add(OutlookThreadContext.Capture());
                context.ExecutionOrder.Add(execution);
                return execution;
            }
            finally
            {
                Interlocked.Decrement(ref context.ActiveExecutions);
            }
        }

        private static string ExecuteStateful(
            DispatcherOwnerContext context,
            string state,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            context.ExecutedThreads.Add(OutlookThreadContext.Capture());
            return state;
        }

        private static int ExecuteStatefulInt(
            DispatcherOwnerContext context,
            int state,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            context.ExecutedThreads.Add(OutlookThreadContext.Capture());
            return state;
        }

        private static int ExecuteBlockingStateful(
            DispatcherOwnerContext context,
            int state,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ExecuteBlockingFirst(context);
            return state;
        }

        private static int ExecuteCooperativelyCanceled(
            DispatcherOwnerContext context,
            int state,
            CancellationToken cancellationToken)
        {
            context.FirstEntered.Set();
            while (!context.ReleaseFirst.IsSet)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Application.DoEvents();
                Thread.Sleep(1);
            }

            return state;
        }

        private static int ThrowStateful(
            DispatcherOwnerContext context,
            int state,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            GC.KeepAlive(context);
            GC.KeepAlive(state);
            throw new InvalidOperationException("stateful failure");
        }

        private static void UpdateMaximum(ref int maximum, int candidate)
        {
            while (true)
            {
                var observed = Volatile.Read(ref maximum);
                if (candidate <= observed ||
                    Interlocked.CompareExchange(ref maximum, candidate, observed) == observed)
                {
                    return;
                }
            }
        }

        private static object[] GetQueuedItems(OutlookStaDispatcher dispatcher)
        {
            var dispatcherType = typeof(OutlookStaDispatcher);
            var gate = dispatcherType
                .GetField("_gate", BindingFlags.Instance | BindingFlags.NonPublic)!
                .GetValue(dispatcher)!;
            var queue = (IEnumerable)dispatcherType
                .GetField("_queue", BindingFlags.Instance | BindingFlags.NonPublic)!
                .GetValue(dispatcher)!;
            lock (gate)
            {
                return queue.Cast<object>().ToArray();
            }
        }

        private static void AssertQueuedContextWorkItemIsIsolated(
            object workItem,
            DispatcherOwnerContext ownerContext)
        {
            Assert.That(workItem.GetType().Name, Does.StartWith("ContextDispatchWorkItem"));
            var fields = workItem.GetType().GetFields(
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            Assert.That(
                fields.Any(field => typeof(DispatcherOwnerContext).IsAssignableFrom(field.FieldType)),
                Is.False);

            foreach (var field in fields)
            {
                var value = field.GetValue(workItem);
                Assert.That(ReferenceEquals(value, ownerContext), Is.False, field.Name);
                Assert.That(value == null || !Marshal.IsComObject(value), Is.True, field.Name);
                if (value is Delegate operation)
                {
                    Assert.That(operation.Target, Is.Null, field.Name);
                }
            }
        }

        private static void AssertFaultedWithMessage(Task task, string expectedMessage)
        {
            Assert.That(
                SpinWait.SpinUntil(() => task.IsCompleted, TimeSpan.FromSeconds(5)),
                Is.True);
            Assert.That(task.IsFaulted, Is.True);
            Assert.That(task.Exception, Is.Not.Null);
            Assert.That(task.Exception!.Flatten().InnerException, Is.TypeOf<InvalidOperationException>());
            Assert.That(task.Exception.Flatten().InnerException!.Message, Is.EqualTo(expectedMessage));
        }

        private sealed class DispatcherOwnerContext : IDisposable
        {
            public readonly List<int> ExecutionOrder = new List<int>();
            public readonly List<OutlookThreadContext> ExecutedThreads =
                new List<OutlookThreadContext>();
            public readonly ManualResetEventSlim FirstEntered = new ManualResetEventSlim();
            public readonly ManualResetEventSlim ReleaseFirst = new ManualResetEventSlim();
            public int ActiveExecutions;
            public int ExecutionCount;
            public int MaximumConcurrentExecutions;

            public int BoundOperation(DispatcherOwnerContext context)
            {
                return ReferenceEquals(this, context) ? context.ExecutionCount : -1;
            }

            public string BoundStateOperation(
                DispatcherOwnerContext context,
                string state,
                CancellationToken cancellationToken)
            {
                cancellationToken.ThrowIfCancellationRequested();
                return ReferenceEquals(this, context) ? state : string.Empty;
            }

            public void Dispose()
            {
                FirstEntered.Dispose();
                ReleaseFirst.Dispose();
            }
        }

        private sealed class StaDispatcherHarness : IDisposable
        {
            private readonly ManualResetEventSlim _ready = new ManualResetEventSlim();
            private readonly ManualResetEventSlim _stop = new ManualResetEventSlim();
            private readonly Thread _thread;
            private readonly DispatcherOwnerContext _context;
            private Exception? _threadFailure;

            public StaDispatcherHarness(DispatcherOwnerContext context)
            {
                _context = context;
                _thread = new Thread(() => Run(context));
                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
                if (!_ready.Wait(TimeSpan.FromSeconds(5)))
                {
                    throw new TimeoutException("The STA dispatcher test thread did not start.");
                }

                if (_threadFailure != null)
                {
                    throw new InvalidOperationException(
                        "The STA dispatcher test thread failed during startup.",
                        _threadFailure);
                }
            }

            public OutlookStaDispatcher Dispatcher { get; private set; } = null!;

            public void Dispose()
            {
                _context.ReleaseFirst.Set();
                _stop.Set();
                Assert.That(_thread.Join(TimeSpan.FromSeconds(5)), Is.True);
                Assert.That(_threadFailure, Is.Null);
                _context.Dispose();
                _ready.Dispose();
                _stop.Dispose();
            }

            private void Run(DispatcherOwnerContext context)
            {
                try
                {
                    Dispatcher = new OutlookStaDispatcher(context);
                    _ready.Set();
                    while (!_stop.IsSet)
                    {
                        Application.DoEvents();
                        Thread.Sleep(1);
                    }

                    Dispatcher.Dispose();
                }
                catch (Exception exception)
                {
                    _threadFailure = exception;
                    _ready.Set();
                }
            }
        }
    }
}
#endif
