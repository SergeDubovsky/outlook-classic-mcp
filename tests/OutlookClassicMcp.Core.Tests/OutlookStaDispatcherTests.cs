#if NET48
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using NUnit.Framework;
using OutlookClassicMcp.AddIn.Runtime;

namespace OutlookClassicMcp.Core.Tests
{
    [TestFixture]
    public sealed class OutlookStaDispatcherTests
    {
        private const int DefaultCapacity = 16;
        private static readonly int[] ExpectedExecutionOrder = { 1, 2 };

        [Test]
        public void NestedMessagePumpDoesNotRunQueuedWorkReentrantly()
        {
            OutlookStaDispatcher? dispatcher = null;
            Exception? uiThreadFailure = null;
            var executionOrder = new List<int>();
            using (var dispatcherReady = new ManualResetEventSlim())
            using (var firstEntered = new ManualResetEventSlim())
            using (var releaseFirst = new ManualResetEventSlim())
            using (var stopUiThread = new ManualResetEventSlim())
            {
                var uiThread = new Thread(() =>
                {
                    try
                    {
                        dispatcher = new OutlookStaDispatcher();
                        dispatcherReady.Set();
                        while (!stopUiThread.IsSet)
                        {
                            Application.DoEvents();
                            Thread.Sleep(1);
                        }

                        dispatcher.Dispose();
                    }
                    catch (Exception exception)
                    {
                        uiThreadFailure = exception;
                        dispatcherReady.Set();
                    }
                });
                uiThread.SetApartmentState(ApartmentState.STA);
                uiThread.Start();

                try
                {
                    Assert.That(dispatcherReady.Wait(TimeSpan.FromSeconds(5)), Is.True);
                    Assert.That(uiThreadFailure, Is.Null);
                    Assert.That(dispatcher, Is.Not.Null);
                    var activeDispatcher = dispatcher
                        ?? throw new InvalidOperationException("The dispatcher was not created.");

                    var first = activeDispatcher.InvokeAsync(
                        () =>
                        {
                            firstEntered.Set();
                            while (!releaseFirst.IsSet)
                            {
                                Application.DoEvents();
                                Thread.Sleep(1);
                            }

                            executionOrder.Add(1);
                            return OutlookThreadContext.Capture();
                        },
                        CancellationToken.None);
                    var second = activeDispatcher.InvokeAsync(
                        () =>
                        {
                            executionOrder.Add(2);
                            return OutlookThreadContext.Capture();
                        },
                        CancellationToken.None);

                    Assert.That(firstEntered.Wait(TimeSpan.FromSeconds(5)), Is.True);
                    Thread.Sleep(100);
                    Assert.That(second.IsCompleted, Is.False);

                    releaseFirst.Set();
                    Assert.That(first.Wait(TimeSpan.FromSeconds(5)), Is.True);
                    Assert.That(second.Wait(TimeSpan.FromSeconds(5)), Is.True);
                    Assert.That(executionOrder, Is.EqualTo(ExpectedExecutionOrder));
                    Assert.That(first.Result.ManagedThreadId, Is.EqualTo(second.Result.ManagedThreadId));
                    Assert.That(first.Result.NativeThreadId, Is.EqualTo(second.Result.NativeThreadId));
                    Assert.That(first.Result.ApartmentState, Is.EqualTo(ApartmentState.STA));
                }
                finally
                {
                    releaseFirst.Set();
                    stopUiThread.Set();
                    Assert.That(uiThread.Join(TimeSpan.FromSeconds(5)), Is.True);
                }

                Assert.That(uiThreadFailure, Is.Null);
            }
        }

        [Test]
        public void BeginShutdownCompletesEveryAcceptedTaskExactlyOnce()
        {
            OutlookStaDispatcher? dispatcher = null;
            Exception? uiThreadFailure = null;
            var queuedOperationExecutions = 0;
            using (var dispatcherReady = new ManualResetEventSlim())
            using (var activeEntered = new ManualResetEventSlim())
            using (var releaseActive = new ManualResetEventSlim())
            using (var stopUiThread = new ManualResetEventSlim())
            using (var queuedCompletions = new CountdownEvent(DefaultCapacity - 1))
            using (var cancellationSource = new CancellationTokenSource())
            {
                var uiThread = new Thread(() =>
                {
                    try
                    {
                        dispatcher = new OutlookStaDispatcher();
                        dispatcherReady.Set();
                        while (!stopUiThread.IsSet)
                        {
                            Application.DoEvents();
                            Thread.Sleep(1);
                        }

                        dispatcher.Dispose();
                    }
                    catch (Exception exception)
                    {
                        uiThreadFailure = exception;
                        dispatcherReady.Set();
                    }
                });
                uiThread.SetApartmentState(ApartmentState.STA);
                uiThread.Start();

                try
                {
                    Assert.That(dispatcherReady.Wait(TimeSpan.FromSeconds(5)), Is.True);
                    Assert.That(uiThreadFailure, Is.Null);
                    Assert.That(dispatcher, Is.Not.Null);
                    var activeDispatcher = dispatcher
                        ?? throw new InvalidOperationException("The dispatcher was not created.");

                    var active = activeDispatcher.InvokeAsync(
                        () =>
                        {
                            activeEntered.Set();
                            releaseActive.Wait();
                            return 42;
                        },
                        CancellationToken.None);
                    var queued = new List<Task<int>>(DefaultCapacity - 1);
                    var completionCounts = new int[DefaultCapacity - 1];
                    for (var index = 0; index < DefaultCapacity - 1; index++)
                    {
                        var capturedIndex = index;
                        var cancellationToken = index == 0
                            ? cancellationSource.Token
                            : CancellationToken.None;
                        var task = activeDispatcher.InvokeAsync(
                            () =>
                            {
                                Interlocked.Increment(ref queuedOperationExecutions);
                                return capturedIndex;
                            },
                            cancellationToken);
                        queued.Add(task);
                        _ = task.ContinueWith(
                            _ =>
                            {
                                Interlocked.Increment(ref completionCounts[capturedIndex]);
                                queuedCompletions.Signal();
                            },
                            CancellationToken.None,
                            TaskContinuationOptions.ExecuteSynchronously,
                            TaskScheduler.Default);
                    }

                    Assert.That(activeEntered.Wait(TimeSpan.FromSeconds(5)), Is.True);
                    Assert.That(activeDispatcher.QueueDepth, Is.EqualTo(DefaultCapacity));

                    var overflow = activeDispatcher.InvokeAsync(() => -1, CancellationToken.None);
                    AssertFaultedWithMessage(overflow, "HOST_BUSY");

                    cancellationSource.Cancel();
                    Assert.That(
                        SpinWait.SpinUntil(() => queued[0].IsCompleted, TimeSpan.FromSeconds(5)),
                        Is.True);
                    Assert.That(queued[0].IsCanceled, Is.True);

                    var firstShutdown = activeDispatcher.BeginShutdown();
                    var secondShutdown = activeDispatcher.BeginShutdown();

                    Assert.That(activeDispatcher.IsAccepting, Is.False);
                    Assert.That(activeDispatcher.QueueDepth, Is.EqualTo(1));
                    Assert.That(secondShutdown, Is.SameAs(firstShutdown));
                    Assert.That(firstShutdown.IsCompleted, Is.False);
                    var rejected = activeDispatcher.InvokeAsync(() => -2, CancellationToken.None);
                    AssertFaultedWithMessage(rejected, "HOST_STOPPING");
                    Assert.That(queuedCompletions.Wait(TimeSpan.FromSeconds(5)), Is.True);

                    for (var index = 1; index < queued.Count; index++)
                    {
                        AssertFaultedWithMessage(queued[index], "HOST_STOPPING");
                    }

                    Assert.That(completionCounts, Is.All.EqualTo(1));
                    Assert.That(queuedOperationExecutions, Is.Zero);

                    releaseActive.Set();
                    var accepted = new List<Task>(queued.Count + 1) { active };
                    accepted.AddRange(queued);
                    Assert.That(
                        SpinWait.SpinUntil(
                            () => accepted.All(task => task.IsCompleted),
                            TimeSpan.FromSeconds(5)),
                        Is.True);
                    Assert.That(active.Status, Is.EqualTo(TaskStatus.RanToCompletion));
                    Assert.That(active.Result, Is.EqualTo(42));
                    Assert.That(firstShutdown.Wait(TimeSpan.FromSeconds(5)), Is.True);
                    Assert.That(completionCounts, Is.All.EqualTo(1));
                }
                finally
                {
                    releaseActive.Set();
                    stopUiThread.Set();
                    Assert.That(uiThread.Join(TimeSpan.FromSeconds(5)), Is.True);
                }

                Assert.That(uiThreadFailure, Is.Null);
            }
        }

        [Test]
        public void WakeupFailureFailsAcceptedWorkAndStopsLifecycleAdmission()
        {
            Exception? uiThreadFailure = null;
            Task<int>? accepted = null;
            var callbackExecutions = 0;
            var firstLifecycleAccepted = true;
            var postAttempts = 0;
            var secondLifecycleAccepted = true;
            var uiThread = new Thread(() =>
            {
                try
                {
                    using (var dispatcher = new OutlookStaDispatcher(
                        (_, _, _, _) =>
                        {
                            Interlocked.Increment(ref postAttempts);
                            return false;
                        }))
                    {
                        accepted = dispatcher.InvokeAsync(() => 1, CancellationToken.None);
                        firstLifecycleAccepted = dispatcher.TryPostLifecycleCallback(
                            () => Interlocked.Increment(ref callbackExecutions));
                        secondLifecycleAccepted = dispatcher.TryPostLifecycleCallback(
                            () => Interlocked.Increment(ref callbackExecutions));
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
            Assert.That(accepted, Is.Not.Null);
            AssertFaultedWithMessage(accepted!, "HOST_UNAVAILABLE");
            Assert.That(firstLifecycleAccepted, Is.False);
            Assert.That(secondLifecycleAccepted, Is.False);
            Assert.That(postAttempts, Is.EqualTo(1));
            Assert.That(callbackExecutions, Is.Zero);
        }

        private static void AssertFaultedWithMessage(Task task, string expectedMessage)
        {
            Assert.That(
                SpinWait.SpinUntil(() => task.IsCompleted, TimeSpan.FromSeconds(5)),
                Is.True);
            Assert.That(task.IsFaulted, Is.True);
            Assert.That(task.Exception, Is.Not.Null);
            Assert.That(task.Exception!.Flatten().InnerExceptions, Has.Count.EqualTo(1));
            Assert.That(task.Exception.Flatten().InnerException, Is.TypeOf<InvalidOperationException>());
            Assert.That(task.Exception.Flatten().InnerException!.Message, Is.EqualTo(expectedMessage));
        }
    }
}
#endif
