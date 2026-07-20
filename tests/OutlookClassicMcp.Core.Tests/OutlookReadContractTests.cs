using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using NUnit.Framework;
using OutlookClassicMcp.Core.Outlook;

namespace OutlookClassicMcp.Core.Tests
{
    [TestFixture]
    public sealed class OutlookReadContractTests
    {
        [Test]
        public void LocatorsAreStoreQualifiedAndExposeNoDisplayNameFallback()
        {
            var mailbox = new MailboxRef("store-1");
            var folder = new FolderRef("store-1", "folder-1");
            var item = new ItemRef("store-1", "item-1", "IPM.Note");

            Assert.Multiple((Action)(() =>
            {
                Assert.That(mailbox.StoreId, Is.EqualTo("store-1"));
                Assert.That(folder.StoreId, Is.EqualTo("store-1"));
                Assert.That(folder.EntryId, Is.EqualTo("folder-1"));
                Assert.That(item.StoreId, Is.EqualTo("store-1"));
                Assert.That(item.EntryId, Is.EqualTo("item-1"));
                Assert.That(item.ItemClass, Is.EqualTo("IPM.Note"));
                Assert.That(typeof(MailboxRef).GetProperty("DisplayName"), Is.Null);
                Assert.That(typeof(FolderRef).GetProperty("DisplayName"), Is.Null);
                Assert.That(typeof(ItemRef).GetProperty("DisplayName"), Is.Null);
            }));
        }

        [TestCase(null)]
        [TestCase("")]
        [TestCase("   ")]
        [TestCase("store id")]
        [TestCase("store\nid")]
        public void OpaqueStoreIdentifiersRejectNonOpaqueValues(string? storeId)
        {
            AssertThrows<ArgumentException>(() => _ = new MailboxRef(storeId!));
        }

        [Test]
        public void OpaqueIdentifierBoundsAreInclusive()
        {
            var maximum = new string('a', MailboxRef.MaximumStoreIdLength);
            var mailbox = new MailboxRef(maximum);

            Assert.That(mailbox.StoreId, Has.Length.EqualTo(MailboxRef.MaximumStoreIdLength));
            AssertThrows<ArgumentOutOfRangeException>(
                () => _ = new MailboxRef(maximum + "a"));
        }

        [Test]
        public void AttachmentLocatorRequiresOneBasedIndexExactFingerprintAndKnownSizeState()
        {
            var item = CreateItem();
            var fingerprint = Fingerprint('a');
            var attachment = new AttachmentRef(
                item,
                1,
                string.Empty,
                0,
                sizeIsKnown: false,
                fingerprint);

            Assert.Multiple((Action)(() =>
            {
                Assert.That(attachment.AttachmentIndex, Is.EqualTo(1));
                Assert.That(attachment.Name, Is.Empty);
                Assert.That(attachment.Size, Is.Zero);
                Assert.That(attachment.SizeIsKnown, Is.False);
                Assert.That(attachment.MetadataFingerprint, Is.EqualTo(fingerprint));
                AssertThrows<ArgumentOutOfRangeException>(
                    () => _ = new AttachmentRef(item, 0, "name", 1, true, fingerprint));
                AssertThrows<ArgumentException>(
                    () => _ = new AttachmentRef(item, 1, "name", 1, false, fingerprint));
                AssertThrows<ArgumentException>(
                    () => _ = new AttachmentRef(item, 1, "name", 1, true, new string('a', 63)));
                AssertThrows<ArgumentException>(
                    () => _ = new AttachmentRef(item, 1, "name", 1, true, new string('A', 64)));
                AssertThrows<ArgumentException>(
                    () => _ = new AttachmentRef(item, 1, "bad\nname", 1, true, fingerprint));
            }));
        }

        [TestCase(1)]
        [TestCase(50)]
        public void EveryPagedRequestAcceptsInclusivePageBounds(int pageSize)
        {
            var item = CreateItem();
            var folder = CreateFolder();
            var filter = EmptyFilter();

            Assert.Multiple((Action)(() =>
            {
                Assert.That(new OutlookListMailboxesRequest(pageSize, null).PageSize, Is.EqualTo(pageSize));
                Assert.That(
                    new OutlookListFoldersRequest(new MailboxRef("store"), null, pageSize, null).PageSize,
                    Is.EqualTo(pageSize));
                Assert.That(new OutlookListMessagesRequest(folder, pageSize, null).PageSize, Is.EqualTo(pageSize));
                Assert.That(
                    new OutlookSearchMessagesRequest(
                        new[] { new OutlookSearchScope(new MailboxRef("store"), null) },
                        filter,
                        pageSize,
                        null).PageSize,
                    Is.EqualTo(pageSize));
                Assert.That(new OutlookGetConversationRequest(item, pageSize, null).PageSize, Is.EqualTo(pageSize));
                Assert.That(new OutlookListAttachmentsRequest(item, pageSize, null).PageSize, Is.EqualTo(pageSize));
            }));
        }

        [TestCase(0)]
        [TestCase(51)]
        public void EveryPagedRequestRejectsValuesOutsidePageBounds(int pageSize)
        {
            var item = CreateItem();
            var folder = CreateFolder();

            Assert.Multiple((Action)(() =>
            {
                AssertThrows<ArgumentOutOfRangeException>(
                    () => _ = new OutlookListMailboxesRequest(pageSize, null));
                AssertThrows<ArgumentOutOfRangeException>(
                    () => _ = new OutlookListFoldersRequest(
                        new MailboxRef("store"),
                        null,
                        pageSize,
                        null));
                AssertThrows<ArgumentOutOfRangeException>(
                    () => _ = new OutlookListMessagesRequest(folder, pageSize, null));
                AssertThrows<ArgumentOutOfRangeException>(
                    () => _ = new OutlookSearchMessagesRequest(
                        new[] { new OutlookSearchScope(new MailboxRef("store"), null) },
                        EmptyFilter(),
                        pageSize,
                        null));
                AssertThrows<ArgumentOutOfRangeException>(
                    () => _ = new OutlookGetConversationRequest(item, pageSize, null));
                AssertThrows<ArgumentOutOfRangeException>(
                    () => _ = new OutlookListAttachmentsRequest(item, pageSize, null));
            }));
        }

        [Test]
        public void RequestsRejectCrossStoreFolderAndMessageAnchors()
        {
            var mailbox = new MailboxRef("store-a");
            var wrongFolder = new FolderRef("store-b", "folder");
            var wrongMessageAnchor = new OutlookMessageKeysetAnchor(
                Utc(1),
                new ItemRef("store-b", "item", "IPM.Note"));

            Assert.Multiple((Action)(() =>
            {
                AssertThrows<ArgumentException>(
                    () => _ = new OutlookListFoldersRequest(mailbox, wrongFolder, 10, null));
                AssertThrows<ArgumentException>(
                    () => _ = new OutlookListMessagesRequest(
                        new FolderRef("store-a", "folder"),
                        10,
                        wrongMessageAnchor));
            }));
        }

        [Test]
        public void ConversationAnchorMayLegitimatelySpanStores()
        {
            var request = new OutlookGetConversationRequest(
                new ItemRef("store-a", "seed", "IPM.Note"),
                10,
                new OutlookMessageKeysetAnchor(
                    Utc(1),
                    new ItemRef("store-b", "page-anchor", "IPM.Note")));

            Assert.That(request.Anchor!.Item.StoreId, Is.EqualTo("store-b"));
        }

        [Test]
        public void NullFolderSearchScopeMeansTheExplicitMailboxInboxScope()
        {
            var mailbox = new MailboxRef("store");
            var scope = new OutlookSearchScope(mailbox, null);

            Assert.Multiple((Action)(() =>
            {
                Assert.That(scope.Mailbox, Is.SameAs(mailbox));
                Assert.That(scope.Folder, Is.Null);
            }));
        }

        [Test]
        public void SearchScopesAreExplicitDistinctBoundedAndDefensivelyCopied()
        {
            var scopes = Enumerable.Range(1, OutlookReadLimits.MaximumSearchScopeCount)
                .Select(index => new OutlookSearchScope(new MailboxRef("store-" + index), null))
                .ToList();
            var first = scopes[0];
            var request = new OutlookSearchMessagesRequest(scopes, EmptyFilter(), 10, null);

            scopes[0] = new OutlookSearchScope(new MailboxRef("replacement"), null);
            scopes.Add(new OutlookSearchScope(new MailboxRef("overflow"), null));

            Assert.Multiple((Action)(() =>
            {
                Assert.That(request.Scopes, Has.Count.EqualTo(64));
                Assert.That(request.Scopes[0], Is.SameAs(first));
                AssertThrows<ArgumentException>(
                    () => _ = new OutlookSearchMessagesRequest(
                        Array.Empty<OutlookSearchScope>(),
                        EmptyFilter(),
                        10,
                        null));
                AssertThrows<ArgumentOutOfRangeException>(
                    () => _ = new OutlookSearchMessagesRequest(scopes, EmptyFilter(), 10, null));
                AssertThrows<ArgumentException>(
                    () => _ = new OutlookSearchMessagesRequest(
                        new[] { first, first },
                        EmptyFilter(),
                        10,
                        null));
            }));
        }

        [Test]
        public void SearchFilterHasOnlyTheFixedTypedFields()
        {
            var properties = typeof(OutlookMessageSearchFilter)
                .GetProperties()
                .Select(property => property.Name)
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray();

            Assert.That(
                properties,
                Is.EqualTo(new[]
                {
                    nameof(OutlookMessageSearchFilter.Category),
                    nameof(OutlookMessageSearchFilter.HasAttachments),
                    nameof(OutlookMessageSearchFilter.IsUnread),
                    nameof(OutlookMessageSearchFilter.ReceivedFromUtc),
                    nameof(OutlookMessageSearchFilter.ReceivedToUtc),
                    nameof(OutlookMessageSearchFilter.Recipient),
                    nameof(OutlookMessageSearchFilter.Sender),
                    nameof(OutlookMessageSearchFilter.Subject),
                    nameof(OutlookMessageSearchFilter.Text),
                }));
        }

        [Test]
        public void SearchFilterRequiresUtcOrderedRangeAndBoundsText()
        {
            Assert.Multiple((Action)(() =>
            {
                AssertThrows<ArgumentException>(
                    () => _ = Filter(receivedFromUtc: DateTime.Now));
                AssertThrows<ArgumentException>(
                    () => _ = Filter(receivedFromUtc: Utc(2), receivedToUtc: Utc(1)));
                AssertThrows<ArgumentOutOfRangeException>(
                    () => _ = Filter(text: new string('x', OutlookMessageSearchFilter.MaximumTextLength + 1)));
                AssertThrows<ArgumentException>(() => _ = Filter(sender: "bad\nsender"));
            }));
        }

        [TestCase(1)]
        [TestCase(50000)]
        public void GetMessageAcceptsInclusiveBodyBounds(int maximumCharacters)
        {
            var request = new OutlookGetMessageRequest(
                CreateItem(),
                OutlookBodyFormat.PlainText,
                maximumCharacters);

            Assert.That(request.MaximumBodyCharacters, Is.EqualTo(maximumCharacters));
        }

        [TestCase(0)]
        [TestCase(50001)]
        public void GetMessageRejectsBodyBoundsOutsideTheLimit(int maximumCharacters)
        {
            AssertThrows<ArgumentOutOfRangeException>(
                () => _ = new OutlookGetMessageRequest(
                    CreateItem(),
                    OutlookBodyFormat.PlainText,
                    maximumCharacters));
        }

        [Test]
        public void BodyAllowsNewlinesAndEnforcesCompleteTruncatedAndProtectedStates()
        {
            var complete = new OutlookMessageBody(
                OutlookBodyFormat.PlainText,
                "line one\r\nline two",
                18,
                isTruncated: false,
                isProtected: false);
            var truncated = new OutlookMessageBody(
                OutlookBodyFormat.Html,
                new string('x', OutlookReadLimits.MaximumBodyCharacters),
                OutlookReadLimits.MaximumBodyCharacters + 1,
                isTruncated: true,
                isProtected: false);
            var protectedBody = new OutlookMessageBody(
                OutlookBodyFormat.PlainText,
                string.Empty,
                null,
                isTruncated: false,
                isProtected: true);

            Assert.Multiple((Action)(() =>
            {
                Assert.That(complete.Content, Does.Contain("\r\n"));
                Assert.That(truncated.IsTruncated, Is.True);
                Assert.That(protectedBody.IsProtected, Is.True);
                AssertThrows<ArgumentException>(
                    () => _ = new OutlookMessageBody(
                        OutlookBodyFormat.PlainText,
                        "secret",
                        null,
                        false,
                        true));
                AssertThrows<ArgumentException>(
                    () => _ = new OutlookMessageBody(
                        OutlookBodyFormat.PlainText,
                        "short",
                        5,
                        true,
                        false));
                AssertThrows<ArgumentException>(
                    () => _ = new OutlookMessageBody(
                        OutlookBodyFormat.PlainText,
                        "short",
                        6,
                        false,
                        false));
                AssertThrows<ArgumentOutOfRangeException>(
                    () => _ = new OutlookMessageBody(
                        OutlookBodyFormat.PlainText,
                        new string('x', OutlookReadLimits.MaximumBodyCharacters + 1),
                        OutlookReadLimits.MaximumBodyCharacters + 1,
                        false,
                        false));
            }));
        }

        [Test]
        public void MessageDetailDefensivelyCopiesBoundedRecipients()
        {
            var first = new OutlookMessageAddress("Person", "person@example.test");
            var recipients = new List<OutlookMessageAddress> { first };
            var detail = new OutlookMessageDetail(
                CreateMessage(1),
                recipients,
                Array.Empty<OutlookMessageAddress>(),
                Array.Empty<OutlookMessageAddress>(),
                totalToRecipientCount: 2,
                totalCcRecipientCount: 0,
                totalBccRecipientCount: 0,
                new OutlookMessageBody(OutlookBodyFormat.PlainText, string.Empty, 0, false, false));

            recipients[0] = new OutlookMessageAddress("Replacement", null);
            recipients.Add(new OutlookMessageAddress("Added", null));

            Assert.Multiple((Action)(() =>
            {
                Assert.That(detail.ToRecipients, Has.Count.EqualTo(1));
                Assert.That(detail.ToRecipients[0], Is.SameAs(first));
                Assert.That(detail.ToRecipientsTruncated, Is.True);
                AssertThrows<ArgumentOutOfRangeException>(
                    () => _ = new OutlookMessageDetail(
                        CreateMessage(1),
                        Enumerable.Range(0, OutlookMessageDetail.MaximumRecipientCount + 1)
                            .Select(index => new OutlookMessageAddress("Recipient " + index, null)),
                        Array.Empty<OutlookMessageAddress>(),
                        Array.Empty<OutlookMessageAddress>(),
                        OutlookMessageDetail.MaximumRecipientCount + 1,
                        0,
                        0,
                        new OutlookMessageBody(
                            OutlookBodyFormat.PlainText,
                            string.Empty,
                            0,
                            false,
                            false)));
            }));
        }

        private static OutlookMessageSearchFilter EmptyFilter()
        {
            return Filter();
        }

        private static OutlookMessageSearchFilter Filter(
            string? sender = null,
            string? text = null,
            DateTime? receivedFromUtc = null,
            DateTime? receivedToUtc = null)
        {
            return new OutlookMessageSearchFilter(
                sender,
                recipient: null,
                subject: null,
                text,
                receivedFromUtc,
                receivedToUtc,
                isUnread: null,
                category: null,
                hasAttachments: null);
        }

        internal static OutlookMessageSummary CreateMessage(int index)
        {
            return new OutlookMessageSummary(
                new ItemRef(
                    "store",
                    "item-" + index.ToString("D4", CultureInfo.InvariantCulture),
                    "IPM.Note"),
                CreateFolder(),
                "Subject " + index,
                "Sender",
                "sender@example.test",
                Utc(2000 - index),
                Utc(2000 - index),
                null,
                isRead: false,
                attachmentCount: 0,
                conversationId: null);
        }

        internal static DateTime Utc(int minute)
        {
            return new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddMinutes(minute);
        }

        internal static ItemRef CreateItem()
        {
            return new ItemRef("store", "item", "IPM.Note");
        }

        internal static FolderRef CreateFolder()
        {
            return new FolderRef("store", "folder");
        }

        internal static string Fingerprint(char value)
        {
            return new string(value, AttachmentRef.FingerprintLength);
        }

        private static void AssertThrows<TException>(Action action)
            where TException : Exception
        {
            Assert.Catch<TException>((Action)(() => action()));
        }
    }
}
