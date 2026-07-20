#if NET48
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using NUnit.Framework;
using OutlookClassicMcp.AddIn.Runtime;
using OutlookClassicMcp.Core.Outlook;

namespace OutlookClassicMcp.Core.Tests
{
    [TestFixture]
    public sealed class OutlookReadImplementationTests
    {
        private static readonly int[] InsertionOrder = { 4, 2, 5, 1, 3 };
        private static readonly string[] ExpectedRetainedItems =
            { "item-0001", "item-0002", "item-0003" };
        private static readonly string[] FirstDeduplicatedPage =
            { "item-01", "item-02" };
        private static readonly string[] SecondDeduplicatedPage =
            { "item-03", "item-04" };

        [Test]
        [NonParallelizable]
        public void SearchRestrictionUsesUtcAndScopeOrderingSchemas()
        {
            var originalCulture = CultureInfo.CurrentCulture;
            try
            {
                CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
                var filter = new OutlookMessageSearchFilter(
                    sender: null,
                    recipient: null,
                    subject: null,
                    text: null,
                    receivedFromUtc: Utc(1, 2, 3),
                    receivedToUtc: Utc(4, 5, 6),
                    isUnread: null,
                    category: null,
                    hasAttachments: null);
                var anchor = new OutlookMessageKeysetAnchor(
                    Utc(7, 8, 9),
                    OutlookReadContractTests.CreateItem());

                var restriction = OutlookReadFilter.BuildSearchRestriction(
                    filter,
                    anchor,
                    OutlookMessageTimestampKind.Sent);

                Assert.Multiple((Action)(() =>
                {
                    Assert.That(
                        restriction,
                        Does.Contain("\"urn:schemas:httpmail:datereceived\" >= '1/15/2026 1:02 AM'"));
                    Assert.That(
                        restriction,
                        Does.Contain("\"urn:schemas:httpmail:datereceived\" <= '1/15/2026 4:06 AM'"));
                    Assert.That(
                        restriction,
                        Does.Contain("\"urn:schemas:httpmail:date\" <= '1/15/2026 7:09 AM'"));
                    Assert.That(restriction, Does.Not.Contain("1/14/2026"));
                }));
            }
            finally
            {
                CultureInfo.CurrentCulture = originalCulture;
            }
        }

        [Test]
        public void SearchRestrictionEscapesLiteralsAndUsesExactCategoryComparison()
        {
            var filter = new OutlookMessageSearchFilter(
                sender: "Sender",
                recipient: null,
                subject: "can't 100%_[x]",
                text: null,
                receivedFromUtc: null,
                receivedToUtc: null,
                isUnread: null,
                category: "Team's",
                hasAttachments: null);

            var restriction = OutlookReadFilter.BuildSearchRestriction(
                filter,
                anchor: null,
                OutlookMessageTimestampKind.Received);

            Assert.Multiple((Action)(() =>
            {
                Assert.That(
                    restriction,
                    Does.Contain("LIKE '%can''t 100[%][_][[]x]%'"));
                Assert.That(
                    restriction,
                    Does.Contain("\"urn:schemas:httpmail:fromemail\" LIKE '%Sender%' OR " +
                        "\"urn:schemas:httpmail:fromname\" LIKE '%Sender%'"));
                Assert.That(
                    restriction,
                    Does.Contain("\"urn:schemas-microsoft-com:office:office#Keywords\" = 'Team''s'"));
            }));
        }

        [Test]
        public void SelectedOrderingTimestampDoesNotFallBackToAnotherField()
        {
            var sent = OutlookReadContractTests.Utc(2);

            Assert.That(
                OutlookReadProjection.SelectEffectiveTimestamp(
                    OutlookMessageTimestampKind.Received,
                    receivedUtc: null,
                    sent,
                    modifiedUtc: null),
                Is.EqualTo(DateTime.SpecifyKind(DateTime.MinValue, DateTimeKind.Utc)));
            Assert.That(
                OutlookReadProjection.SelectEffectiveTimestamp(
                    OutlookMessageTimestampKind.Automatic,
                    receivedUtc: null,
                    sent,
                    modifiedUtc: null),
                Is.EqualTo(sent));
        }

        [Test]
        public void BoundedInsertionRetainsOnlyTheHighestRankedMessages()
        {
            var candidates = new List<OutlookMessageSummary>();
            foreach (var index in InsertionOrder)
            {
                OutlookReadProjection.InsertBounded(
                    candidates,
                    OutlookReadContractTests.CreateMessage(index),
                    maximumCount: 3,
                    OutlookReadProjection.CompareMessages);
            }

            Assert.That(
                candidates.Select(value => value.Item.EntryId).ToArray(),
                Is.EqualTo(ExpectedRetainedItems));
        }

        [Test]
        public void MessageOrderingUsesStoreAndEntryIdentifiersForTimestampTies()
        {
            var timestamp = OutlookReadContractTests.Utc(2);
            var first = CreateMessage("store-a", "item-b", timestamp);
            var second = CreateMessage("store-b", "item-a", timestamp);
            var sameStoreLaterItem = CreateMessage("store-a", "item-c", timestamp);

            Assert.Multiple((Action)(() =>
            {
                Assert.That(
                    OutlookReadProjection.CompareMessages(first, second),
                    Is.LessThan(0));
                Assert.That(
                    OutlookReadProjection.CompareMessages(first, sameStoreLaterItem),
                    Is.LessThan(0));
                Assert.That(
                    OutlookReadProjection.CompareMessageToAnchor(
                        first,
                        new OutlookMessageKeysetAnchor(timestamp, second.Item)),
                    Is.EqualTo(OutlookReadProjection.CompareMessages(first, second)));
            }));
        }

        [Test]
        public void EqualTimestampKeysetsRemainContinuousAcrossThreeProductionPages()
        {
            var timestamp = OutlookReadContractTests.Utc(2);
            var source = new List<OutlookMessageSummary>
            {
                CreateMessage("store-b", "item-04", timestamp),
                CreateMessage("store-a", "item-03", timestamp),
                CreateMessage("store-b", "item-01", timestamp),
                CreateMessage("store-a", "item-05", timestamp),
                CreateMessage("store-a", "item-01", timestamp),
                CreateMessage("store-b", "item-03", timestamp),
                CreateMessage("store-a", "item-04", timestamp),
                CreateMessage("store-b", "item-02", timestamp),
                CreateMessage("store-a", "item-02", timestamp),
            };
            var expected = new List<OutlookMessageSummary>(source);
            expected.Sort(OutlookReadProjection.CompareMessages);
            var returned = new List<OutlookMessageSummary>();
            OutlookMessageKeysetAnchor? anchor = null;

            for (var pageNumber = 1; pageNumber <= 3; pageNumber++)
            {
                var candidates = source
                    .Where(value => anchor == null ||
                        OutlookReadProjection.CompareMessageToAnchor(value, anchor) > 0)
                    .ToList();
                var page = OutlookReadProjection.BuildMessagePage(
                    candidates,
                    pageSize: 3,
                    totalScopeCount: 1,
                    Array.Empty<OutlookScopeFailure>());

                Assert.That(page.Items, Has.Count.EqualTo(3));
                Assert.That(
                    page.NextAnchor,
                    pageNumber < 3 ? Is.Not.Null : Is.Null);
                returned.AddRange(page.Items);
                anchor = page.NextAnchor;
            }

            Assert.That(
                returned.Select(value => value.Item.StoreId + "/" + value.Item.EntryId),
                Is.EqualTo(expected.Select(
                    value => value.Item.StoreId + "/" + value.Item.EntryId)));
        }

        [Test]
        public void MailboxAndFolderComparatorsAreSymmetricWithTheirAnchors()
        {
            var firstMailbox = CreateMailbox("Alpha", "store-a");
            var secondMailbox = CreateMailbox("Alpha", "store-b");
            var firstFolder = new OutlookFolderSummary(
                new FolderRef("store-a", "folder-a"),
                parentFolder: null,
                "Folder",
                hasChildren: false);
            var secondFolder = new OutlookFolderSummary(
                new FolderRef("store-a", "folder-b"),
                parentFolder: null,
                "Folder",
                hasChildren: false);

            Assert.Multiple((Action)(() =>
            {
                Assert.That(
                    OutlookReadProjection.CompareMailboxToAnchor(
                        firstMailbox,
                        new OutlookMailboxKeysetAnchor(
                            secondMailbox.DisplayName,
                            secondMailbox.Mailbox)),
                    Is.EqualTo(OutlookReadProjection.CompareMailboxes(
                        firstMailbox,
                        secondMailbox)));
                Assert.That(
                    OutlookReadProjection.CompareFolderToAnchor(
                        firstFolder,
                        new OutlookFolderKeysetAnchor(
                            secondFolder.DisplayName,
                            secondFolder.Folder)),
                    Is.EqualTo(OutlookReadProjection.CompareFolders(
                        firstFolder,
                        secondFolder)));
            }));
        }

        [Test]
        public void OptionalUtcConversionEnforcesUsefulBoundsAndDateTimeKind()
        {
            var utc = new DateTime(2026, 7, 19, 12, 30, 0, DateTimeKind.Utc);
            var local = new DateTime(2026, 7, 19, 12, 30, 0, DateTimeKind.Local);
            var unspecified = new DateTime(
                2026,
                7,
                19,
                12,
                30,
                0,
                DateTimeKind.Unspecified);

            Assert.Multiple((Action)(() =>
            {
                Assert.That(
                    OutlookReadProjection.ToOptionalUtc(new DateTime(1899, 12, 31)),
                    Is.Null);
                Assert.That(
                    OutlookReadProjection.ToOptionalUtc(new DateTime(4500, 1, 1)),
                    Is.Null);
                Assert.That(OutlookReadProjection.ToOptionalUtc(utc), Is.EqualTo(utc));
                Assert.That(
                    OutlookReadProjection.ToOptionalUtc(local),
                    Is.EqualTo(local.ToUniversalTime()));
                Assert.That(
                    OutlookReadProjection.ToOptionalUtc(unspecified),
                    Is.EqualTo(DateTime.SpecifyKind(
                        unspecified,
                        DateTimeKind.Local).ToUniversalTime()));
            }));
        }

        [Test]
        public void SearchMergeRecordsMappedScopedFailureAlongsideSuccess()
        {
            var successfulScope = new OutlookSearchScope(
                new MailboxRef("store-a"),
                folder: null);
            var failedScope = new OutlookSearchScope(
                new MailboxRef("store-b"),
                folder: null);
            var merge = new OutlookSearchMerge(pageSize: 2, totalScopeCount: 2);
            merge.AddSuccess(new[]
            {
                CreateMessage("store-a", "item-01", OutlookReadContractTests.Utc(2)),
            });
            merge.AddFailure(
                failedScope,
                Marshal.GetExceptionForHR(unchecked((int)0x80010001)) ??
                    throw new InvalidOperationException("The test HRESULT was not mapped."));

            var page = merge.Complete();

            Assert.Multiple((Action)(() =>
            {
                Assert.That(page.Items, Has.Count.EqualTo(1));
                Assert.That(page.IsPartial, Is.True);
                Assert.That(page.Failures, Has.Count.EqualTo(1));
                Assert.That(page.Failures[0].Scope, Is.SameAs(failedScope));
                Assert.That(
                    page.Failures[0].Failure,
                    Is.EqualTo(OutlookGatewayFailure.ComBusy));
                Assert.That(page.TotalScopeCount, Is.EqualTo(2));
                Assert.That(successfulScope.Mailbox.StoreId, Is.EqualTo("store-a"));
            }));
        }

        [Test]
        public void SearchMergeThrowsTheFirstTypedFailureWhenEveryScopeFails()
        {
            var merge = new OutlookSearchMerge(pageSize: 2, totalScopeCount: 2);
            merge.AddFailure(
                new OutlookSearchScope(new MailboxRef("store-a"), folder: null),
                new OutlookGatewayException(OutlookGatewayFailure.Timeout));
            merge.AddFailure(
                new OutlookSearchScope(new MailboxRef("store-b"), folder: null),
                new OutlookGatewayException(OutlookGatewayFailure.AccessDenied));

            Action complete = () => merge.Complete();

            Assert.That(
                complete,
                Throws.TypeOf<OutlookGatewayException>()
                    .With.Property(nameof(OutlookGatewayException.Failure))
                    .EqualTo(OutlookGatewayFailure.Timeout));
        }

        [Test]
        public void SearchMergeSuppressesCursorWhenAnOverflowingPageIsPartial()
        {
            var merge = new OutlookSearchMerge(pageSize: 2, totalScopeCount: 2);
            merge.AddSuccess(new[]
            {
                CreateMessage("store-a", "item-01", OutlookReadContractTests.Utc(3)),
                CreateMessage("store-a", "item-02", OutlookReadContractTests.Utc(2)),
                CreateMessage("store-a", "item-03", OutlookReadContractTests.Utc(1)),
            });
            merge.AddFailure(
                new OutlookSearchScope(new MailboxRef("store-b"), folder: null),
                new OutlookGatewayException(OutlookGatewayFailure.UnsupportedStore));

            var page = merge.Complete();

            Assert.Multiple((Action)(() =>
            {
                Assert.That(page.Items, Has.Count.EqualTo(2));
                Assert.That(page.IsPartial, Is.True);
                Assert.That(page.NextAnchor, Is.Null);
            }));
        }

        [Test]
        public void SearchMergeDeduplicatesOverlappingScopesAcrossContinuation()
        {
            var firstItem = CreateMessage(
                "store-a",
                "item-01",
                OutlookReadContractTests.Utc(4));
            var secondItem = CreateMessage(
                "store-a",
                "item-02",
                OutlookReadContractTests.Utc(3));
            var thirdItem = CreateMessage(
                "store-a",
                "item-03",
                OutlookReadContractTests.Utc(2));
            var fourthItem = CreateMessage(
                "store-a",
                "item-04",
                OutlookReadContractTests.Utc(1));
            var firstMerge = new OutlookSearchMerge(pageSize: 2, totalScopeCount: 2);
            firstMerge.AddSuccess(new[] { firstItem, secondItem, thirdItem });
            firstMerge.AddSuccess(new[] { firstItem, secondItem, thirdItem });

            var firstPage = firstMerge.Complete();
            var continuation = firstPage.NextAnchor;
            Assert.That(continuation, Is.Not.Null);
            var secondMerge = new OutlookSearchMerge(pageSize: 2, totalScopeCount: 2);
            secondMerge.AddSuccess(new[] { thirdItem, fourthItem });
            secondMerge.AddSuccess(new[] { thirdItem, fourthItem });
            var secondPage = secondMerge.Complete();

            Assert.Multiple((Action)(() =>
            {
                Assert.That(
                    firstPage.Items.Select(value => value.Item.EntryId),
                    Is.EqualTo(FirstDeduplicatedPage));
                Assert.That(
                    firstPage.Items
                        .Select(value => value.Item.EntryId)
                        .Distinct()
                        .Count(),
                    Is.EqualTo(2));
                Assert.That(
                    OutlookReadProjection.CompareMessageToAnchor(
                        thirdItem,
                        continuation!),
                    Is.GreaterThan(0));
                Assert.That(
                    secondPage.Items.Select(value => value.Item.EntryId),
                    Is.EqualTo(SecondDeduplicatedPage));
                Assert.That(
                    secondPage.Items
                        .Select(value => value.Item.EntryId)
                        .Distinct()
                        .Count(),
                    Is.EqualTo(2));
                Assert.That(secondPage.NextAnchor, Is.Null);
                Assert.That(firstPage.TotalScopeCount, Is.EqualTo(2));
                Assert.That(secondPage.TotalScopeCount, Is.EqualTo(2));
            }));
        }

        [Test]
        public void CanonicalScopeOrderingMakesDuplicateSelectionIndependentOfInputOrder()
        {
            var implicitInbox = new OutlookSearchScope(
                new MailboxRef("store-a"),
                folder: null);
            var explicitInbox = new OutlookSearchScope(
                new MailboxRef("store-a"),
                new FolderRef("store-a", "inbox-folder"));
            var implicitProjection = CreateMessage(
                "store-a",
                "same-item",
                OutlookReadContractTests.Utc(2));
            var explicitProjection = CreateMessage(
                "store-a",
                "same-item",
                OutlookReadContractTests.Utc(3));

            var forward = MergeCanonicalScopes(
                new[] { implicitInbox, explicitInbox },
                implicitProjection,
                explicitProjection);
            var reversed = MergeCanonicalScopes(
                new[] { explicitInbox, implicitInbox },
                implicitProjection,
                explicitProjection);

            Assert.Multiple((Action)(() =>
            {
                Assert.That(forward.Items, Has.Count.EqualTo(1));
                Assert.That(reversed.Items, Has.Count.EqualTo(1));
                Assert.That(
                    forward.Items[0].EffectiveTimestampUtc,
                    Is.EqualTo(implicitProjection.EffectiveTimestampUtc));
                Assert.That(
                    reversed.Items[0].EffectiveTimestampUtc,
                    Is.EqualTo(forward.Items[0].EffectiveTimestampUtc));
                Assert.That(forward.TotalScopeCount, Is.EqualTo(2));
                Assert.That(reversed.TotalScopeCount, Is.EqualTo(2));
            }));
        }

        [Test]
        public void SearchFolderPolicyIsBoundedAndRejectsDetectedMatches()
        {
            Action acceptBound = () => OutlookSearchFolderPolicy.RequireBoundedCount(
                OutlookSearchFolderPolicy.MaximumSearchFoldersExamined);
            Action rejectOverflow = () => OutlookSearchFolderPolicy.RequireBoundedCount(
                OutlookSearchFolderPolicy.MaximumSearchFoldersExamined + 1);
            Action acceptNormalFolder = () => OutlookSearchFolderPolicy.RejectIfMatch(false);
            Action rejectSearchFolder = () => OutlookSearchFolderPolicy.RejectIfMatch(true);

            Assert.Multiple((Action)(() =>
            {
                Assert.That(acceptBound, Throws.Nothing);
                Assert.That(
                    rejectOverflow,
                    Throws.TypeOf<OutlookGatewayException>()
                        .With.Property(nameof(OutlookGatewayException.Failure))
                        .EqualTo(OutlookGatewayFailure.Timeout));
                Assert.That(acceptNormalFolder, Throws.Nothing);
                Assert.That(
                    rejectSearchFolder,
                    Throws.TypeOf<OutlookGatewayException>()
                        .With.Property(nameof(OutlookGatewayException.Failure))
                        .EqualTo(OutlookGatewayFailure.UnsupportedStore));
            }));
        }

        [Test]
        public void PartialMergedPageSuppressesTheGlobalContinuationAnchor()
        {
            var candidates = Enumerable.Range(1, 3)
                .Select(OutlookReadContractTests.CreateMessage)
                .ToList();
            var failedScope = new OutlookSearchScope(new MailboxRef("failed-store"), null);

            var partial = OutlookReadProjection.BuildMessagePage(
                new List<OutlookMessageSummary>(candidates),
                pageSize: 2,
                totalScopeCount: 2,
                new[]
                {
                    new OutlookScopeFailure(failedScope, OutlookGatewayFailure.Timeout),
                });
            var complete = OutlookReadProjection.BuildMessagePage(
                candidates,
                pageSize: 2,
                totalScopeCount: 1,
                Array.Empty<OutlookScopeFailure>());

            Assert.Multiple((Action)(() =>
            {
                Assert.That(partial.Items, Has.Count.EqualTo(2));
                Assert.That(partial.IsPartial, Is.True);
                Assert.That(partial.NextAnchor, Is.Null);
                Assert.That(complete.NextAnchor, Is.Not.Null);
            }));
        }

        [Test]
        public void ProjectionBoundsDisplayTextButNeverRewritesOpaqueIdentifiers()
        {
            Assert.That(OutlookReadProjection.BoundDisplay("a\r\nbcdef", 4), Is.EqualTo("abcd"));
            Action rejectWhitespace = () =>
                OutlookReadProjection.RequireIdentifier("item id", 20);
            Action rejectOverflow = () =>
                OutlookReadProjection.RequireIdentifier("12345", 4);
            Assert.That(
                rejectWhitespace,
                Throws.TypeOf<OutlookGatewayException>()
                    .With.Property(nameof(OutlookGatewayException.Failure))
                    .EqualTo(OutlookGatewayFailure.Internal));
            Assert.That(
                rejectOverflow,
                Throws.TypeOf<OutlookGatewayException>());
        }

        [Test]
        public void AttachmentFingerprintIsDeterministicAndMetadataSensitive()
        {
            var item = OutlookReadContractTests.CreateItem();
            var first = OutlookReadProjection.ComputeAttachmentFingerprint(
                item,
                1,
                "file.txt",
                "file.txt",
                10,
                1,
                0,
                0,
                "text/plain");
            var same = OutlookReadProjection.ComputeAttachmentFingerprint(
                item,
                1,
                "file.txt",
                "file.txt",
                10,
                1,
                0,
                0,
                "text/plain");
            var changed = OutlookReadProjection.ComputeAttachmentFingerprint(
                item,
                1,
                "file.txt",
                "file.txt",
                11,
                1,
                0,
                0,
                "text/plain");

            Assert.Multiple((Action)(() =>
            {
                Assert.That(first, Is.EqualTo(same));
                Assert.That(first, Is.Not.EqualTo(changed));
                Assert.That(first, Does.Match("^[0-9a-f]{64}$"));
            }));
        }

        private static DateTime Utc(int hour, int minute, int second)
        {
            return new DateTime(2026, 1, 15, hour, minute, second, DateTimeKind.Utc);
        }

        private static OutlookMessageSummary CreateMessage(
            string storeId,
            string entryId,
            DateTime effectiveTimestampUtc)
        {
            return new OutlookMessageSummary(
                new ItemRef(storeId, entryId, "IPM.Note"),
                new FolderRef(storeId, "folder"),
                subject: string.Empty,
                senderDisplayName: null,
                senderAddress: null,
                effectiveTimestampUtc,
                receivedUtc: effectiveTimestampUtc,
                sentUtc: null,
                isRead: true,
                attachmentCount: 0,
                conversationId: null);
        }

        private static OutlookMailboxSummary CreateMailbox(string displayName, string storeId)
        {
            return new OutlookMailboxSummary(
                new MailboxRef(storeId),
                displayName,
                OutlookStoreType.NonExchange,
                new OutlookStoreCapabilities(
                    isExchangeStore: false,
                    isDataFileStore: true,
                    isCachedExchange: false),
                new OutlookStandardFolderReferences(null, null, null, null, null));
        }

        private static OutlookMessagePage MergeCanonicalScopes(
            OutlookSearchScope[] scopes,
            OutlookMessageSummary implicitProjection,
            OutlookMessageSummary explicitProjection)
        {
            var merge = new OutlookSearchMerge(pageSize: 1, totalScopeCount: scopes.Length);
            foreach (var scope in OutlookSearchMerge.CanonicalizeScopes(scopes))
            {
                merge.AddSuccess(new[]
                {
                    scope.Folder == null ? implicitProjection : explicitProjection,
                });
            }

            return merge.Complete();
        }
    }
}
#endif
