using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OutlookClassicMcp.Core.Outlook;

namespace OutlookClassicMcp.Core.Tests
{
    [TestFixture]
    public sealed class OutlookProbeContractTests
    {
        [TestCase(32)]
        [TestCase(64)]
        public void SnapshotAcceptsSupportedBitness(int bitness)
        {
            var snapshot = CreateSnapshot(outlookBitness: bitness);

            Assert.That(snapshot.OutlookBitness, Is.EqualTo(bitness));
        }

        [TestCase(0)]
        [TestCase(31)]
        [TestCase(86)]
        [TestCase(128)]
        public void SnapshotRejectsUnsupportedBitness(int bitness)
        {
            AssertThrows<ArgumentOutOfRangeException>(
                () => CreateSnapshot(outlookBitness: bitness));
        }

        [Test]
        public void SnapshotPreservesManagedProbeData()
        {
            var store = CreateStore(
                displayName: "Primary Mailbox",
                storeType: OutlookStoreType.PrimaryExchangeMailbox);
            var thread = CreateThreadProof();
            var snapshot = new OutlookProbeSnapshot(
                "16.0.12345.10000",
                64,
                "Test Profile",
                thread,
                1,
                new[] { store },
                new[] { OutlookProbeWarning.ArchiveNotExposedByOutlookObjectModel });

            AssertAll(() =>
            {
                Assert.That(snapshot.OutlookVersion, Is.EqualTo("16.0.12345.10000"));
                Assert.That(snapshot.ActiveProfileName, Is.EqualTo("Test Profile"));
                Assert.That(snapshot.DispatcherThread, Is.SameAs(thread));
                Assert.That(snapshot.ConfiguredStoreCount, Is.EqualTo(1));
                Assert.That(snapshot.Stores, Is.EqualTo(new[] { store }));
                Assert.That(
                    snapshot.Warnings,
                    Is.EqualTo(new[] { OutlookProbeWarning.ArchiveNotExposedByOutlookObjectModel }));
                Assert.That(snapshot.IsPartial, Is.True);
            });
        }

        [Test]
        public void SnapshotCopiesStoreAndWarningCollections()
        {
            var originalStore = CreateStore();
            var stores = new List<OutlookStoreProbe> { originalStore };
            var warnings = new List<OutlookProbeWarning>
            {
                OutlookProbeWarning.ArchiveNotExposedByOutlookObjectModel,
            };
            var snapshot = CreateSnapshot(
                configuredStoreCount: 1,
                stores: stores,
                warnings: warnings);

            stores[0] = CreateStore(displayName: "Replacement");
            stores.Add(CreateStore(displayName: "Added"));
            warnings[0] = OutlookProbeWarning.StoreMetadataIncomplete;
            warnings.Add(OutlookProbeWarning.StoreLimitReached);

            AssertAll(() =>
            {
                Assert.That(snapshot.Stores, Has.Count.EqualTo(1));
                Assert.That(snapshot.Stores[0], Is.SameAs(originalStore));
                Assert.That(snapshot.Warnings, Has.Count.EqualTo(1));
                Assert.That(
                    snapshot.Warnings[0],
                    Is.EqualTo(OutlookProbeWarning.ArchiveNotExposedByOutlookObjectModel));
            });
        }

        [Test]
        public void SnapshotExposesReadOnlyCollections()
        {
            var snapshot = CreateSnapshot(
                configuredStoreCount: 1,
                stores: new[] { CreateStore() },
                warnings: new[] { OutlookProbeWarning.StoreMetadataIncomplete });
            var stores = (IList<OutlookStoreProbe>)snapshot.Stores;
            var warnings = (IList<OutlookProbeWarning>)snapshot.Warnings;

            AssertAll(() =>
            {
                AssertActionThrows<NotSupportedException>(() => stores.Add(CreateStore()));
                AssertActionThrows<NotSupportedException>(
                    () => warnings.Add(OutlookProbeWarning.StoreLimitReached));
            });
        }

        [Test]
        public void SnapshotAcceptsAtMostSixtyFourReturnedStores()
        {
            var stores = Enumerable.Range(1, OutlookProbeSnapshot.MaximumStoreCount)
                .Select(index => CreateStore(displayName: "Store " + index))
                .ToArray();

            var snapshot = CreateSnapshot(
                configuredStoreCount: stores.Length,
                stores: stores);

            Assert.That(snapshot.Stores, Has.Count.EqualTo(OutlookProbeSnapshot.MaximumStoreCount));
        }

        [Test]
        public void SnapshotRejectsMoreThanSixtyFourReturnedStores()
        {
            var stores = Enumerable.Range(1, OutlookProbeSnapshot.MaximumStoreCount + 1)
                .Select(index => CreateStore(displayName: "Store " + index))
                .ToArray();

            AssertThrows<ArgumentOutOfRangeException>(
                () => CreateSnapshot(configuredStoreCount: stores.Length, stores: stores));
        }

        [Test]
        public void SnapshotAllowsConfiguredCountBeyondReturnedBound()
        {
            var stores = Enumerable.Range(1, OutlookProbeSnapshot.MaximumStoreCount)
                .Select(index => CreateStore(displayName: "Store " + index))
                .ToArray();

            var snapshot = CreateSnapshot(
                configuredStoreCount: OutlookProbeSnapshot.MaximumStoreCount + 10,
                stores: stores,
                warnings: new[] { OutlookProbeWarning.StoreLimitReached });

            AssertAll(() =>
            {
                Assert.That(
                    snapshot.ConfiguredStoreCount,
                    Is.EqualTo(OutlookProbeSnapshot.MaximumStoreCount + 10));
                Assert.That(snapshot.Stores, Has.Count.EqualTo(OutlookProbeSnapshot.MaximumStoreCount));
                Assert.That(snapshot.IsPartial, Is.True);
            });
        }

        [Test]
        public void SnapshotRejectsConfiguredCountBelowReturnedCount()
        {
            AssertThrows<ArgumentOutOfRangeException>(
                () => CreateSnapshot(
                    configuredStoreCount: 0,
                    stores: new[] { CreateStore() }));
        }

        [Test]
        public void SnapshotRejectsNegativeConfiguredCount()
        {
            AssertThrows<ArgumentOutOfRangeException>(
                () => CreateSnapshot(configuredStoreCount: -1));
        }

        [Test]
        public void SnapshotRejectsNullStoreAndWarningCollections()
        {
            AssertAll(() =>
            {
                AssertThrows<ArgumentNullException>(
                    () => new OutlookProbeSnapshot(
                        "16.0",
                        64,
                        "Profile",
                        CreateThreadProof(),
                        0,
                        null!,
                        Array.Empty<OutlookProbeWarning>()));
                AssertThrows<ArgumentNullException>(
                    () => new OutlookProbeSnapshot(
                        "16.0",
                        64,
                        "Profile",
                        CreateThreadProof(),
                        0,
                        Array.Empty<OutlookStoreProbe>(),
                        null!));
            });
        }

        [Test]
        public void SnapshotRejectsNullStoreElements()
        {
            AssertThrows<ArgumentException>(
                () => CreateSnapshot(
                    configuredStoreCount: 1,
                    stores: new OutlookStoreProbe[] { null! }));
        }

        [Test]
        public void SnapshotBoundsWarningsAndRejectsUndefinedWarnings()
        {
            var maximumWarnings = Enumerable.Repeat(
                    OutlookProbeWarning.StoreMetadataIncomplete,
                    OutlookProbeSnapshot.MaximumWarningCount)
                .ToArray();
            var tooManyWarnings = maximumWarnings
                .Concat(new[] { OutlookProbeWarning.StoreLimitReached })
                .ToArray();

            AssertAll(() =>
            {
                Assert.That(
                    CreateSnapshot(warnings: maximumWarnings).Warnings,
                    Has.Count.EqualTo(OutlookProbeSnapshot.MaximumWarningCount));
                AssertThrows<ArgumentOutOfRangeException>(
                    () => CreateSnapshot(warnings: tooManyWarnings));
                AssertThrows<ArgumentOutOfRangeException>(
                    () => CreateSnapshot(warnings: new[] { (OutlookProbeWarning)99 }));
            });
        }

        [Test]
        public void SnapshotDerivesPartialStateFromMissingStoresOrWarnings()
        {
            var complete = CreateSnapshot();
            var missingStore = CreateSnapshot(configuredStoreCount: 1);
            var warning = CreateSnapshot(
                warnings: new[] { OutlookProbeWarning.ArchiveNotExposedByOutlookObjectModel });

            AssertAll(() =>
            {
                Assert.That(complete.IsPartial, Is.False);
                Assert.That(missingStore.IsPartial, Is.True);
                Assert.That(warning.IsPartial, Is.True);
            });
        }

        [Test]
        public void VersionProfileAndDisplayNameAcceptTheirMaximumLengths()
        {
            var version = new string('v', OutlookProbeSnapshot.MaximumVersionLength);
            var profile = new string('p', OutlookProbeSnapshot.MaximumProfileNameLength);
            var displayName = new string('d', OutlookStoreProbe.MaximumDisplayNameLength);
            var store = CreateStore(displayName: displayName);

            var snapshot = CreateSnapshot(
                outlookVersion: version,
                activeProfileName: profile,
                configuredStoreCount: 1,
                stores: new[] { store });

            AssertAll(() =>
            {
                Assert.That(snapshot.OutlookVersion, Has.Length.EqualTo(64));
                Assert.That(snapshot.ActiveProfileName, Has.Length.EqualTo(256));
                Assert.That(snapshot.Stores[0].DisplayName, Has.Length.EqualTo(256));
            });
        }

        [Test]
        public void VersionProfileAndDisplayNameRejectValuesBeyondTheirBounds()
        {
            AssertAll(() =>
            {
                AssertThrows<ArgumentOutOfRangeException>(
                    () => CreateSnapshot(
                        outlookVersion: new string('v', OutlookProbeSnapshot.MaximumVersionLength + 1)));
                AssertThrows<ArgumentOutOfRangeException>(
                    () => CreateSnapshot(
                        activeProfileName: new string('p', OutlookProbeSnapshot.MaximumProfileNameLength + 1)));
                AssertThrows<ArgumentOutOfRangeException>(
                    () => CreateStore(
                        displayName: new string('d', OutlookStoreProbe.MaximumDisplayNameLength + 1)));
            });
        }

        [TestCase(null)]
        [TestCase("")]
        [TestCase("   ")]
        [TestCase("bad\rvalue")]
        [TestCase("bad\u0000value")]
        public void VersionRejectsNullEmptyWhitespaceAndControlText(string? value)
        {
            AssertThrows<ArgumentException>(() => CreateSnapshot(outlookVersion: value!));
        }

        [TestCase(null)]
        [TestCase("")]
        [TestCase("   ")]
        [TestCase("bad\nvalue")]
        [TestCase("bad\u001fvalue")]
        public void ProfileRejectsNullEmptyWhitespaceAndControlText(string? value)
        {
            AssertThrows<ArgumentException>(() => CreateSnapshot(activeProfileName: value!));
        }

        [TestCase(null)]
        [TestCase("")]
        [TestCase("   ")]
        [TestCase("bad\tvalue")]
        [TestCase("bad\u007fvalue")]
        public void DisplayNameRejectsNullEmptyWhitespaceAndControlText(string? value)
        {
            AssertThrows<ArgumentException>(() => CreateStore(displayName: value!));
        }

        [Test]
        public void StoreProbePreservesTypeCapabilitiesAndFolderAvailability()
        {
            var capabilities = new OutlookStoreCapabilities(
                isExchangeStore: true,
                isDataFileStore: false,
                isCachedExchange: true);
            var folders = new StandardFolderAvailability(
                OutlookFolderAvailability.Available,
                OutlookFolderAvailability.Missing,
                OutlookFolderAvailability.Available,
                OutlookFolderAvailability.Unknown,
                OutlookFolderAvailability.Unknown);
            var store = new OutlookStoreProbe(
                "Shared Mailbox",
                OutlookStoreType.AdditionalExchangeMailbox,
                capabilities,
                folders);

            AssertAll(() =>
            {
                Assert.That(store.StoreType, Is.EqualTo(OutlookStoreType.AdditionalExchangeMailbox));
                Assert.That(store.Capabilities, Is.SameAs(capabilities));
                Assert.That(store.StandardFolders, Is.SameAs(folders));
                Assert.That(store.Capabilities.IsExchangeStore, Is.True);
                Assert.That(store.Capabilities.IsDataFileStore, Is.False);
                Assert.That(store.Capabilities.IsCachedExchange, Is.True);
                Assert.That(store.StandardFolders.Inbox, Is.EqualTo(OutlookFolderAvailability.Available));
                Assert.That(store.StandardFolders.Drafts, Is.EqualTo(OutlookFolderAvailability.Missing));
                Assert.That(store.StandardFolders.Sent, Is.EqualTo(OutlookFolderAvailability.Available));
                Assert.That(store.StandardFolders.Deleted, Is.EqualTo(OutlookFolderAvailability.Unknown));
                Assert.That(store.StandardFolders.Archive, Is.EqualTo(OutlookFolderAvailability.Unknown));
            });
        }

        [Test]
        public void StoreProbeRejectsUndefinedTypeAndNullNestedContracts()
        {
            AssertAll(() =>
            {
                AssertThrows<ArgumentOutOfRangeException>(
                    () => CreateStore(storeType: (OutlookStoreType)99));
                AssertThrows<ArgumentNullException>(
                    () => new OutlookStoreProbe(
                        "Store",
                        OutlookStoreType.Unknown,
                        null!,
                        CreateFolders()));
                AssertThrows<ArgumentNullException>(
                    () => new OutlookStoreProbe(
                        "Store",
                        OutlookStoreType.Unknown,
                        new OutlookStoreCapabilities(false, false, false),
                        null!));
            });
        }

        [TestCase(0)]
        [TestCase(1)]
        [TestCase(2)]
        [TestCase(3)]
        [TestCase(4)]
        public void StandardFoldersRejectUndefinedAvailability(int folderIndex)
        {
            var values = Enumerable.Repeat(OutlookFolderAvailability.Available, 5).ToArray();
            values[folderIndex] = (OutlookFolderAvailability)99;

            AssertThrows<ArgumentOutOfRangeException>(
                () => new StandardFolderAvailability(
                    values[0],
                    values[1],
                    values[2],
                    values[3],
                    values[4]));
        }

        [Test]
        public void ThreadProofPreservesMatchingManagedAndFullRangeNativeIds()
        {
            var proof = new OutlookDispatcherThreadProof(
                capturedManagedThreadId: 17,
                capturedNativeThreadId: uint.MaxValue,
                executedManagedThreadId: 17,
                executedNativeThreadId: uint.MaxValue,
                executedOnSta: true);

            AssertAll(() =>
            {
                Assert.That(proof.CapturedManagedThreadId, Is.EqualTo(17));
                Assert.That(proof.CapturedNativeThreadId, Is.EqualTo(uint.MaxValue));
                Assert.That(proof.ExecutedManagedThreadId, Is.EqualTo(17));
                Assert.That(proof.ExecutedNativeThreadId, Is.EqualTo(uint.MaxValue));
                Assert.That(proof.ExecutedOnSta, Is.True);
            });
        }

        [Test]
        public void ThreadProofRejectsNonPositiveThreadIds()
        {
            AssertAll(() =>
            {
                AssertThrows<ArgumentOutOfRangeException>(
                    () => new OutlookDispatcherThreadProof(0, 2, 0, 2, true));
                AssertThrows<ArgumentOutOfRangeException>(
                    () => new OutlookDispatcherThreadProof(-1, 2, -1, 2, true));
                AssertThrows<ArgumentOutOfRangeException>(
                    () => new OutlookDispatcherThreadProof(1, 0, 1, 0, true));
                AssertThrows<ArgumentOutOfRangeException>(
                    () => new OutlookDispatcherThreadProof(1, 2, 0, 2, true));
                AssertThrows<ArgumentOutOfRangeException>(
                    () => new OutlookDispatcherThreadProof(1, 2, 1, 0, true));
            });
        }

        [Test]
        public void ThreadProofRejectsAThreadMismatchOrNonStaExecution()
        {
            AssertAll(() =>
            {
                AssertThrows<ArgumentException>(
                    () => new OutlookDispatcherThreadProof(1, 2, 3, 2, true));
                AssertThrows<ArgumentException>(
                    () => new OutlookDispatcherThreadProof(1, 2, 1, 3, true));
                AssertThrows<ArgumentException>(
                    () => new OutlookDispatcherThreadProof(1, 2, 1, 2, false));
            });
        }

        private static OutlookProbeSnapshot CreateSnapshot(
            string outlookVersion = "16.0",
            int outlookBitness = 64,
            string activeProfileName = "Profile",
            int configuredStoreCount = 0,
            IEnumerable<OutlookStoreProbe>? stores = null,
            IEnumerable<OutlookProbeWarning>? warnings = null)
        {
            return new OutlookProbeSnapshot(
                outlookVersion,
                outlookBitness,
                activeProfileName,
                CreateThreadProof(),
                configuredStoreCount,
                stores ?? Array.Empty<OutlookStoreProbe>(),
                warnings ?? Array.Empty<OutlookProbeWarning>());
        }

        private static OutlookDispatcherThreadProof CreateThreadProof()
        {
            return new OutlookDispatcherThreadProof(11, 12, 11, 12, true);
        }

        private static OutlookStoreProbe CreateStore(
            string displayName = "Mailbox",
            OutlookStoreType storeType = OutlookStoreType.Unknown)
        {
            return new OutlookStoreProbe(
                displayName,
                storeType,
                new OutlookStoreCapabilities(false, false, false),
                CreateFolders());
        }

        private static StandardFolderAvailability CreateFolders()
        {
            return new StandardFolderAvailability(
                OutlookFolderAvailability.Available,
                OutlookFolderAvailability.Available,
                OutlookFolderAvailability.Available,
                OutlookFolderAvailability.Available,
                OutlookFolderAvailability.Unknown);
        }

        private static void AssertAll(Action assertions)
        {
            Assert.Multiple((Action)(() => assertions()));
        }

        private static void AssertThrows<TException>(Func<object?> action)
            where TException : Exception
        {
            Assert.Catch<TException>((Action)(() => _ = action()));
        }

        private static void AssertActionThrows<TException>(Action action)
            where TException : Exception
        {
            Assert.Catch<TException>((Action)(() => action()));
        }

    }
}
