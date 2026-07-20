using System;
using System.Globalization;
using System.Linq;
using NUnit.Framework;
using OutlookClassicMcp.Core.Outlook;

namespace OutlookClassicMcp.Core.Tests
{
    [TestFixture]
    public sealed class OutlookReadPaginationTests
    {
        [Test]
        public void ThreeConsecutiveKeysetPagesFromLargeFolderHaveNoGapsOrDuplicates()
        {
            var source = Enumerable.Range(1, 1205)
                .Select(OutlookReadContractTests.CreateMessage)
                .ToArray();

            var first = ReadPage(source, null, 50);
            var second = ReadPage(source, first.NextAnchor, 50);
            var third = ReadPage(source, second.NextAnchor, 50);
            var returnedIds = first.Items
                .Concat(second.Items)
                .Concat(third.Items)
                .Select(message => message.Item.EntryId)
                .ToArray();

            Assert.Multiple((Action)(() =>
            {
                Assert.That(source, Has.Length.GreaterThan(1000));
                Assert.That(returnedIds, Has.Length.EqualTo(150));
                Assert.That(
                    returnedIds.Distinct(StringComparer.Ordinal).ToArray(),
                    Has.Length.EqualTo(150));
                Assert.That(
                    returnedIds,
                    Is.EqualTo(Enumerable.Range(1, 150).Select(
                        index => "item-" + index.ToString("D4", CultureInfo.InvariantCulture))));
                Assert.That(first.NextAnchor, Is.Not.Null);
                Assert.That(second.NextAnchor, Is.Not.Null);
                Assert.That(third.NextAnchor, Is.Not.Null);
            }));
        }

        [Test]
        public void MessagePageRequiresAnchorToMatchItsFinalItem()
        {
            var items = new[]
            {
                OutlookReadContractTests.CreateMessage(1),
                OutlookReadContractTests.CreateMessage(2),
            };

            Assert.Multiple((Action)(() =>
            {
                AssertThrows<ArgumentException>(
                    () => _ = new OutlookMessagePage(
                        items,
                        AnchorFor(items[0]),
                        1,
                        Array.Empty<OutlookScopeFailure>()));
                AssertThrows<ArgumentException>(
                    () => _ = new OutlookMessagePage(
                        Array.Empty<OutlookMessageSummary>(),
                        AnchorFor(items[0]),
                        1,
                        Array.Empty<OutlookScopeFailure>()));
                AssertThrows<ArgumentException>(
                    () => _ = new OutlookMessagePage(
                        items,
                        new OutlookMessageKeysetAnchor(
                            items[1].EffectiveTimestampUtc,
                            new ItemRef(
                                items[1].Item.StoreId,
                                items[1].Item.EntryId,
                                "IPM.Post")),
                        1,
                        Array.Empty<OutlookScopeFailure>()));
                Assert.That(
                    new OutlookMessagePage(
                        items,
                        AnchorFor(items[1]),
                        1,
                        Array.Empty<OutlookScopeFailure>()).NextAnchor,
                    Is.Not.Null);
            }));
        }

        [Test]
        public void EveryPageRejectsMoreThanFiftyItemsAndDefensivelyCopiesItems()
        {
            var messages = Enumerable.Range(1, OutlookReadLimits.MaximumPageSize)
                .Select(OutlookReadContractTests.CreateMessage)
                .ToList();
            var first = messages[0];
            var page = new OutlookMessagePage(
                messages,
                null,
                1,
                Array.Empty<OutlookScopeFailure>());

            messages[0] = OutlookReadContractTests.CreateMessage(999);
            messages.Add(OutlookReadContractTests.CreateMessage(1000));

            Assert.Multiple((Action)(() =>
            {
                Assert.That(page.Items, Has.Count.EqualTo(50));
                Assert.That(page.Items[0], Is.SameAs(first));
                AssertThrows<ArgumentOutOfRangeException>(
                    () => _ = new OutlookMessagePage(
                        messages,
                        null,
                        1,
                        Array.Empty<OutlookScopeFailure>()));
            }));
        }

        [Test]
        public void CrossScopePartialStateRequiresAtLeastOneSuccessAndUniqueScopedFailures()
        {
            var failedScope = new OutlookSearchScope(new MailboxRef("failed-store"), null);
            var failure = new OutlookScopeFailure(
                failedScope,
                OutlookGatewayFailure.AccessDenied);
            var item = OutlookReadContractTests.CreateMessage(1);
            var partial = new OutlookMessagePage(
                Array.Empty<OutlookMessageSummary>(),
                null,
                totalScopeCount: 2,
                new[] { failure });
            var complete = new OutlookMessagePage(
                Array.Empty<OutlookMessageSummary>(),
                null,
                totalScopeCount: 2,
                Array.Empty<OutlookScopeFailure>());

            Assert.Multiple((Action)(() =>
            {
                Assert.That(partial.IsPartial, Is.True);
                Assert.That(partial.Failures, Is.EqualTo(new[] { failure }));
                Assert.That(complete.IsPartial, Is.False);
                AssertThrows<ArgumentException>(
                    () => _ = new OutlookMessagePage(
                        Array.Empty<OutlookMessageSummary>(),
                        null,
                        totalScopeCount: 1,
                        new[] { failure }));
                AssertThrows<ArgumentException>(
                    () => _ = new OutlookMessagePage(
                        Array.Empty<OutlookMessageSummary>(),
                        null,
                        totalScopeCount: 3,
                        new[] { failure, failure }));
                AssertThrows<ArgumentException>(
                    () => _ = new OutlookMessagePage(
                        new[] { item },
                        AnchorFor(item),
                        totalScopeCount: 2,
                        new[] { failure }));
                AssertThrows<ArgumentException>(
                    () => _ = new OutlookScopeFailure(
                        failedScope,
                        OutlookGatewayFailure.Internal));
            }));
        }

        [Test]
        public void MailboxFolderAndAttachmentPagesRequireTypedFinalAnchors()
        {
            var mailbox = new OutlookMailboxSummary(
                new MailboxRef("store"),
                string.Empty,
                OutlookStoreType.NonExchange,
                new OutlookStoreCapabilities(false, true, false),
                new OutlookStandardFolderReferences(null, null, null, null, null));
            var folder = new OutlookFolderSummary(
                new FolderRef("store", "folder"),
                null,
                string.Empty,
                hasChildren: false);
            var attachmentRef = new AttachmentRef(
                OutlookReadContractTests.CreateItem(),
                1,
                string.Empty,
                0,
                sizeIsKnown: false,
                OutlookReadContractTests.Fingerprint('b'));
            var attachment = new OutlookAttachmentSummary(attachmentRef, null);

            Assert.Multiple((Action)(() =>
            {
                Assert.That(
                    new OutlookMailboxPage(
                        new[] { mailbox },
                        new OutlookMailboxKeysetAnchor(string.Empty, mailbox.Mailbox)).NextAnchor,
                    Is.Not.Null);
                Assert.That(
                    new OutlookFolderPage(
                        new[] { folder },
                        new OutlookFolderKeysetAnchor(string.Empty, folder.Folder)).NextAnchor,
                    Is.Not.Null);
                Assert.That(
                    new OutlookAttachmentPage(
                        new[] { attachment },
                        new OutlookAttachmentKeysetAnchor(
                            attachmentRef.AttachmentIndex,
                            attachmentRef.MetadataFingerprint)).NextAnchor,
                    Is.Not.Null);
                AssertThrows<ArgumentException>(
                    () => _ = new OutlookMailboxPage(
                        Array.Empty<OutlookMailboxSummary>(),
                        new OutlookMailboxKeysetAnchor(string.Empty, mailbox.Mailbox)));
                AssertThrows<ArgumentException>(
                    () => _ = new OutlookFolderPage(
                        new[] { folder },
                        new OutlookFolderKeysetAnchor(
                            "different",
                            folder.Folder)));
                AssertThrows<ArgumentException>(
                    () => _ = new OutlookAttachmentPage(
                        new[] { attachment },
                        new OutlookAttachmentKeysetAnchor(
                            2,
                            attachmentRef.MetadataFingerprint)));
            }));
        }

        private static OutlookMessagePage ReadPage(
            OutlookMessageSummary[] source,
            OutlookMessageKeysetAnchor? anchor,
            int pageSize)
        {
            var start = 0;
            if (anchor != null)
            {
                while (start < source.Length && !Matches(source[start], anchor))
                {
                    start++;
                }

                if (start == source.Length)
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.CursorStale);
                }

                start++;
            }

            var count = Math.Min(pageSize, source.Length - start);
            var items = new OutlookMessageSummary[count];
            for (var index = 0; index < count; index++)
            {
                items[index] = source[start + index];
            }

            var hasMore = start + count < source.Length;
            var nextAnchor = hasMore && count > 0 ? AnchorFor(items[count - 1]) : null;
            return new OutlookMessagePage(
                items,
                nextAnchor,
                1,
                Array.Empty<OutlookScopeFailure>());
        }

        private static OutlookMessageKeysetAnchor AnchorFor(OutlookMessageSummary message)
        {
            return new OutlookMessageKeysetAnchor(message.EffectiveTimestampUtc, message.Item);
        }

        private static bool Matches(
            OutlookMessageSummary message,
            OutlookMessageKeysetAnchor anchor)
        {
            return message.EffectiveTimestampUtc == anchor.EffectiveTimestampUtc &&
                string.Equals(message.Item.StoreId, anchor.Item.StoreId, StringComparison.Ordinal) &&
                string.Equals(message.Item.EntryId, anchor.Item.EntryId, StringComparison.Ordinal);
        }

        private static void AssertThrows<TException>(Action action)
            where TException : Exception
        {
            Assert.Catch<TException>((Action)(() => action()));
        }
    }
}
