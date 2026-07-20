using System;
using System.Threading;
using System.Threading.Tasks;
using NUnit.Framework;
using OutlookClassicMcp.AddIn.Runtime;

namespace OutlookClassicMcp.Core.Tests
{
    [TestFixture]
    [NonParallelizable]
    public sealed class OutlookComMetricsTests
    {
        [SetUp]
        public void SetUp()
        {
            OutlookComMetrics.ResetForTests();
        }

        [TearDown]
        public void TearDown()
        {
            OutlookComMetrics.ResetForTests();
        }

        [Test]
        public void TracksCumulativeOwnershipAndPeakWithoutReceivingComObjects()
        {
            OutlookComMetrics.RecordAcquired();
            OutlookComMetrics.RecordAcquired();
            OutlookComMetrics.RecordReleased();
            OutlookComMetrics.RecordReleased();

            var snapshot = OutlookComMetrics.Capture();

            Assert.Multiple((Action)(() =>
            {
                Assert.That(snapshot.ComAcquired, Is.EqualTo(2));
                Assert.That(snapshot.ComReleased, Is.EqualTo(2));
                Assert.That(snapshot.ComOutstanding, Is.Zero);
                Assert.That(snapshot.ComPeak, Is.EqualTo(2));
            }));
        }

        [Test]
        public void MaterializedItemHighWaterOnlyIncreases()
        {
            OutlookComMetrics.ObserveMaterializedItems(4);
            OutlookComMetrics.ObserveMaterializedItems(2);
            OutlookComMetrics.ObserveMaterializedItems(7);

            Assert.That(
                OutlookComMetrics.Capture().MaterializedItemHighWater,
                Is.EqualTo(7));
        }

        [Test]
        public void ConcurrentCaptureNeverObservesATornOwnershipSnapshot()
        {
            const int iterations = 10_000;
            using (var readerReady = new ManualResetEventSlim(false))
            using (var start = new ManualResetEventSlim(false))
            {
                var writerFinished = 0;
                var captureCount = 0;
                var reader = Task.Run(() =>
                {
                    readerReady.Set();
                    start.Wait();
                    while (Volatile.Read(ref writerFinished) == 0)
                    {
                        var snapshot = OutlookComMetrics.Capture();
                        if (snapshot.ComOutstanding !=
                            snapshot.ComAcquired - snapshot.ComReleased)
                        {
                            throw new InvalidOperationException(
                                "The captured ownership counters were torn.");
                        }

                        Interlocked.Increment(ref captureCount);
                    }
                });
                var writer = Task.Run(() =>
                {
                    start.Wait();
                    try
                    {
                        for (var index = 0; index < iterations; index++)
                        {
                            OutlookComMetrics.RecordAcquired();
                            if ((index & 31) == 0)
                            {
                                Thread.Yield();
                            }

                            OutlookComMetrics.RecordReleased();
                        }
                    }
                    finally
                    {
                        Volatile.Write(ref writerFinished, 1);
                    }
                });

                readerReady.Wait();
                start.Set();
                Task.WaitAll(writer, reader);

                var final = OutlookComMetrics.Capture();
                Assert.Multiple((Action)(() =>
                {
                    Assert.That(captureCount, Is.GreaterThan(0));
                    Assert.That(final.ComAcquired, Is.EqualTo(iterations));
                    Assert.That(final.ComReleased, Is.EqualTo(iterations));
                    Assert.That(final.ComOutstanding, Is.Zero);
                }));
            }
        }
    }
}
