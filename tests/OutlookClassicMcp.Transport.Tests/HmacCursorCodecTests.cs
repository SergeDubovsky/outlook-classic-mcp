using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;

namespace OutlookClassicMcp.Transport.Tests
{
    [TestFixture]
    public sealed class HmacCursorCodecTests
    {
        private const string TokenText =
            "AAECAwQFBgcICQoLDA0ODxAREhMUFRYXGBkaGxwdHh8";
        private const string ReplacementTokenText =
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA";
        private const string KeyDomain = "outlook-classic-mcp/cursor-key/v1";
        private const string QueryHash =
            "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
        private const string OtherQueryHash =
            "bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb";
        private const string Fingerprint =
            "cccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc";
        private const long TimestampUtcTicks = 639174672000000000L;

        private static readonly int[] AllKinds =
        {
            (int)HmacCursorKind.ListMailboxes,
            (int)HmacCursorKind.ListFolders,
            (int)HmacCursorKind.ListMessages,
            (int)HmacCursorKind.SearchMessages,
            (int)HmacCursorKind.GetConversation,
            (int)HmacCursorKind.ListAttachments,
        };

        private static readonly string[] InvalidAuthenticatedMailboxPayloads =
        {
            "{",
            " {\"v\":1,\"kind\":\"listMailboxes\",\"queryHash\":\"" + QueryHash +
                "\",\"displayName\":\"Mailbox\",\"storeId\":\"STORE-ID\"}",
            "{\"v\":1,\"kind\":\"listMailboxes\",\"queryHash\":\"" + QueryHash +
                "\",\"displayName\":\"Mailbox\",\"storeId\":\"STORE-ID\",\"extra\":true}",
            "{\"v\":1,\"kind\":\"listMailboxes\",\"queryHash\":\"" + QueryHash +
                "\",\"displayName\":\"Mailbox\"}",
            "{\"v\":1,\"kind\":\"listMailboxes\",\"queryHash\":\"" + QueryHash +
                "\",\"displayName\":\"Mailbox\",\"displayName\":\"Mailbox\"," +
                "\"storeId\":\"STORE-ID\"}",
            "{\"kind\":\"listMailboxes\",\"v\":1,\"queryHash\":\"" + QueryHash +
                "\",\"displayName\":\"Mailbox\",\"storeId\":\"STORE-ID\"}",
            "{\"v\":2,\"kind\":\"listMailboxes\",\"queryHash\":\"" + QueryHash +
                "\",\"displayName\":\"Mailbox\",\"storeId\":\"STORE-ID\"}",
            "{\"v\":1,\"kind\":\"listFolders\",\"queryHash\":\"" + QueryHash +
                "\",\"displayName\":\"Mailbox\",\"storeId\":\"STORE-ID\"}",
            "{\"v\":1,\"kind\":\"listMailboxes\",\"queryHash\":\"" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
                "\",\"displayName\":\"Mailbox\",\"storeId\":\"STORE-ID\"}",
            "{\"v\":1,\"kind\":\"listMailboxes\",\"queryHash\":\"" + QueryHash +
                "\",\"displayName\":null,\"storeId\":\"STORE-ID\"}",
            "{\"v\":1,\"kind\":\"listMailboxes\",\"queryHash\":\"" + QueryHash +
                "\",\"displayName\":\"\\u004dailbox\",\"storeId\":\"STORE-ID\"}",
            "{\"v\":1,\"kind\":\"listMailboxes\",\"queryHash\":\"" + QueryHash +
                "\",\"displayName\":\"Line\\nBreak\",\"storeId\":\"STORE-ID\"}",
        };

        [TestCaseSource(nameof(AllKinds))]
        public void RoundTripsEveryTypedCursorKind(int kindValue)
        {
            var kind = (HmacCursorKind)kindValue;
            using (var codec = CreateCodec())
            {
                var source = CreatePayload(kind);
                var cursor = codec.Encode(source);

                Assert.That(cursor, Does.StartWith("v1."));
                Assert.That(cursor.Length, Is.LessThanOrEqualTo(HmacCursorCodec.MaximumCursorLength));
                Assert.That(
                    codec.TryDecode(cursor, kind, QueryHash, out var decoded),
                    Is.True);
                AssertPayloadEqual(source, decoded);
            }
        }

        [Test]
        public void EncodingIsDeterministicCanonicalAndUsesTheFixedDerivedKeyDomain()
        {
            var payload = new MailboxCursorPayload(QueryHash, "Mailbox", "STORE-ID");
            using (var first = CreateCodec())
            using (var second = CreateCodec())
            {
                var firstCursor = first.Encode(payload);
                var secondCursor = first.Encode(payload);
                var independentCursor = second.Encode(payload);
                const string canonicalJson =
                    "{\"v\":1,\"kind\":\"listMailboxes\",\"queryHash\":\"" +
                    QueryHash +
                    "\",\"displayName\":\"Mailbox\",\"storeId\":\"STORE-ID\"}";

                Assert.That(secondCursor, Is.EqualTo(firstCursor));
                Assert.That(independentCursor, Is.EqualTo(firstCursor));
                Assert.That(firstCursor, Does.Not.Contain(TokenText));
                Assert.That(DecodePayloadJson(firstCursor), Is.EqualTo(canonicalJson));
                Assert.That(SignPayloadJson(canonicalJson, TokenText), Is.EqualTo(firstCursor));
            }
        }

        [Test]
        public void TamperTruncationBadVersionSegmentsAndOversizeAreIndistinguishable()
        {
            using (var codec = CreateCodec())
            {
                var cursor = codec.Encode(CreatePayload(HmacCursorKind.ListMailboxes));
                var separators = FindSeparators(cursor);
                var invalidValues = new string?[]
                {
                    null,
                    string.Empty,
                    "v1",
                    "v1..",
                    "v2" + cursor.Substring(2),
                    cursor + ".extra",
                    cursor.Substring(0, cursor.Length - 1),
                    ReplaceCharacter(cursor, separators.Item1 + 1),
                    ReplaceCharacter(cursor, separators.Item2 + 1),
                    cursor.Substring(0, separators.Item1 + 1) + "=" +
                        cursor.Substring(separators.Item1 + 2),
                    new string('A', HmacCursorCodec.MaximumCursorLength + 1),
                };

                foreach (var invalid in invalidValues)
                {
                    Assert.That(
                        codec.TryDecode(
                            invalid,
                            HmacCursorKind.ListMailboxes,
                            QueryHash,
                            out _),
                        Is.False,
                        invalid == null ? "null" : invalid.Substring(0, Math.Min(32, invalid.Length)));
                }
            }
        }

        [Test]
        public void RejectsAuthenticatedNonCanonicalBase64ForPayloadAndTag()
        {
            using (var codec = CreateCodec())
            {
                var payload = new MailboxCursorPayload(QueryHash, "Mailbox", "STORE-ID");
                var canonicalCursor = codec.Encode(payload);
                var segments = canonicalCursor.Split('.');
                var nonCanonicalPayload = FindNonCanonicalEquivalent(segments[1]);
                var signedNonCanonicalPayload = SignPayloadSegment(
                    nonCanonicalPayload,
                    TokenText);
                var nonCanonicalTag = FindNonCanonicalEquivalent(segments[2]);
                var cursorWithNonCanonicalTag = segments[0] + "." + segments[1] + "." +
                    nonCanonicalTag;

                Assert.That(
                    codec.TryDecode(
                        signedNonCanonicalPayload,
                        HmacCursorKind.ListMailboxes,
                        QueryHash,
                        out _),
                    Is.False);
                Assert.That(
                    codec.TryDecode(
                        cursorWithNonCanonicalTag,
                        HmacCursorKind.ListMailboxes,
                        QueryHash,
                        out _),
                    Is.False);
            }
        }

        [TestCaseSource(nameof(InvalidAuthenticatedMailboxPayloads))]
        public void AuthenticatedJsonMustBeClosedBoundedAndCanonical(string json)
        {
            using (var codec = CreateCodec())
            {
                var cursor = SignPayloadJson(json, TokenText);

                Assert.That(
                    codec.TryDecode(
                        cursor,
                        HmacCursorKind.ListMailboxes,
                        QueryHash,
                        out _),
                    Is.False);
            }
        }

        [TestCaseSource(nameof(GetInvalidTypedPayloadCases))]
        public void AuthenticatedTypedPayloadBoundsAreRejected(int kindValue, string json)
        {
            var kind = (HmacCursorKind)kindValue;
            using (var codec = CreateCodec())
            {
                var cursor = SignPayloadJson(json, TokenText);

                Assert.That(
                    codec.TryDecode(cursor, kind, QueryHash, out _),
                    Is.False);
            }
        }

        [Test]
        public void WrongKindQueryHashAndRotatedKeyReturnTheSameInvalidResult()
        {
            using (var original = CreateCodec())
            using (var replacement = CreateCodec(ReplacementTokenText))
            {
                var cursor = original.Encode(CreatePayload(HmacCursorKind.ListMessages));

                Assert.That(
                    original.TryDecode(
                        cursor,
                        HmacCursorKind.SearchMessages,
                        QueryHash,
                        out _),
                    Is.False);
                Assert.That(
                    original.TryDecode(
                        cursor,
                        HmacCursorKind.ListMessages,
                        OtherQueryHash,
                        out _),
                    Is.False);
                Assert.That(
                    original.TryDecode(
                        cursor,
                        HmacCursorKind.ListMessages,
                        QueryHash.ToUpperInvariant(),
                        out _),
                    Is.False);
                Assert.That(
                    original.TryDecode(
                        cursor,
                        (HmacCursorKind)99,
                        QueryHash,
                        out _),
                    Is.False);
                Assert.That(
                    replacement.TryDecode(
                        cursor,
                        HmacCursorKind.ListMessages,
                        QueryHash,
                        out _),
                    Is.False);
            }
        }

        [Test]
        public void ConstructorsAndEncodingEnforceEveryBound()
        {
            Assert.Throws<ArgumentException>((Action)(() =>
                _ = new MailboxCursorPayload(QueryHash.Substring(1), "Mailbox", "STORE")));
            Assert.Throws<ArgumentException>((Action)(() =>
                _ = new MailboxCursorPayload(QueryHash.ToUpperInvariant(), "Mailbox", "STORE")));
            Assert.DoesNotThrow((Action)(() =>
                _ = new MailboxCursorPayload(QueryHash, string.Empty, "STORE")));
            Assert.Throws<ArgumentException>((Action)(() =>
                _ = new MailboxCursorPayload(QueryHash, "Line\nBreak", "STORE")));
            Assert.Throws<ArgumentOutOfRangeException>((Action)(() =>
                _ = new MailboxCursorPayload(
                    QueryHash,
                    new string('d', HmacCursorCodec.MaximumDisplayNameLength + 1),
                    "STORE")));
            Assert.Throws<ArgumentException>((Action)(() =>
                _ = new MailboxCursorPayload(QueryHash, "Mailbox", string.Empty)));
            Assert.Throws<ArgumentException>((Action)(() =>
                _ = new MailboxCursorPayload(QueryHash, "Mailbox", "STORE ID")));
            Assert.Throws<ArgumentException>((Action)(() =>
                _ = new MailboxCursorPayload(QueryHash, "Mailbox", "STORE\nID")));
            Assert.Throws<ArgumentOutOfRangeException>((Action)(() =>
                _ = new MailboxCursorPayload(
                    QueryHash,
                    "Mailbox",
                    new string('s', HmacCursorCodec.MaximumOpaqueIdentifierLength + 1))));
            Assert.Throws<ArgumentOutOfRangeException>((Action)(() =>
                _ = new MessageCursorPayload(
                    HmacCursorKind.ListMessages,
                    QueryHash,
                    -1,
                    "STORE",
                    "ENTRY",
                    "IPM.Note")));
            Assert.Throws<ArgumentOutOfRangeException>((Action)(() =>
                _ = new MessageCursorPayload(
                    HmacCursorKind.ListMessages,
                    QueryHash,
                    DateTime.MaxValue.Ticks + 1,
                    "STORE",
                    "ENTRY",
                    "IPM.Note")));
            Assert.Throws<ArgumentOutOfRangeException>((Action)(() =>
                _ = new MessageCursorPayload(
                    HmacCursorKind.ListFolders,
                    QueryHash,
                    TimestampUtcTicks,
                    "STORE",
                    "ENTRY",
                    "IPM.Note")));
            Assert.DoesNotThrow((Action)(() =>
                _ = new MessageCursorPayload(
                    HmacCursorKind.ListMessages,
                    QueryHash,
                    DateTime.MaxValue.Ticks,
                    "STORE",
                    "ENTRY",
                    "IPM. Note")));
            Assert.Throws<ArgumentException>((Action)(() =>
                _ = new MessageCursorPayload(
                    HmacCursorKind.ListMessages,
                    QueryHash,
                    TimestampUtcTicks,
                    "STORE",
                    "ENTRY",
                    " ")));
            Assert.Throws<ArgumentOutOfRangeException>((Action)(() =>
                _ = new MessageCursorPayload(
                    HmacCursorKind.ListMessages,
                    QueryHash,
                    TimestampUtcTicks,
                    "STORE",
                    "ENTRY",
                    new string('i', HmacCursorCodec.MaximumItemClassLength + 1))));
            Assert.Throws<ArgumentOutOfRangeException>((Action)(() =>
                _ = new AttachmentCursorPayload(QueryHash, 0, Fingerprint)));
            Assert.Throws<ArgumentException>((Action)(() =>
                _ = new AttachmentCursorPayload(
                    QueryHash,
                    1,
                    Fingerprint.ToUpperInvariant())));

            using (var codec = CreateCodec())
            {
                var maximumMailbox = new MailboxCursorPayload(
                    QueryHash,
                    new string('\uffff', HmacCursorCodec.MaximumDisplayNameLength),
                    new string('s', HmacCursorCodec.MaximumOpaqueIdentifierLength));
                var maximumFolder = new FolderCursorPayload(
                    QueryHash,
                    new string('\uffff', HmacCursorCodec.MaximumDisplayNameLength),
                    new string('s', HmacCursorCodec.MaximumOpaqueIdentifierLength),
                    new string('e', HmacCursorCodec.MaximumOpaqueIdentifierLength));
                var maximumMessage = new MessageCursorPayload(
                    HmacCursorKind.ListMessages,
                    QueryHash,
                    DateTime.MaxValue.Ticks,
                    new string('s', HmacCursorCodec.MaximumOpaqueIdentifierLength),
                    new string('e', HmacCursorCodec.MaximumOpaqueIdentifierLength),
                    new string('\uffff', HmacCursorCodec.MaximumItemClassLength));

                foreach (var maximumPayload in new HmacCursorPayload[]
                {
                    maximumMailbox,
                    maximumFolder,
                    maximumMessage,
                })
                {
                    var maximumCursor = codec.Encode(maximumPayload);
                    Assert.That(maximumCursor, Has.Length.LessThanOrEqualTo(
                        HmacCursorCodec.MaximumCursorLength));
                    Assert.That(
                        codec.TryDecode(
                            maximumCursor,
                            maximumPayload.Kind,
                            QueryHash,
                            out var maximumDecoded),
                        Is.True);
                    Assert.That(maximumDecoded.Kind, Is.EqualTo(maximumPayload.Kind));
                }

                var maximumAttachment = new AttachmentCursorPayload(
                    QueryHash,
                    int.MaxValue,
                    Fingerprint);
                var cursor = codec.Encode(maximumAttachment);
                Assert.That(
                    codec.TryDecode(
                        cursor,
                        HmacCursorKind.ListAttachments,
                        QueryHash,
                        out var decoded),
                    Is.True);
                Assert.That(
                    ((AttachmentCursorPayload)decoded).AttachmentIndex,
                    Is.EqualTo(int.MaxValue));
            }
        }

        [Test]
        public void DisposedCodecFailsClosedAndDisposeIsIdempotent()
        {
            var codec = CreateCodec();
            var payload = CreatePayload(HmacCursorKind.ListMailboxes);
            var cursor = codec.Encode(payload);
            codec.Dispose();
            Assert.DoesNotThrow((Action)codec.Dispose);

            Assert.Throws<ObjectDisposedException>((Action)(() => codec.Encode(payload)));
            Assert.Throws<ObjectDisposedException>((Action)(() =>
                codec.TryDecode(
                    cursor,
                    HmacCursorKind.ListMailboxes,
                    QueryHash,
                    out _)));
        }

        [Test]
        public async Task ConcurrentEncodingAndDecodingAreDeterministicAndThreadSafe()
        {
            using (var codec = CreateCodec())
            {
                var tasks = Enumerable.Range(0, 64)
                    .Select(index => Task.Run(() =>
                    {
                        var payload = new MessageCursorPayload(
                            HmacCursorKind.SearchMessages,
                            QueryHash,
                            TimestampUtcTicks + index,
                            "STORE-" + index,
                            "ENTRY-" + index,
                            "IPM.Note");
                        var first = codec.Encode(payload);
                        var second = codec.Encode(payload);
                        return string.Equals(first, second, StringComparison.Ordinal) &&
                            codec.TryDecode(
                                first,
                                HmacCursorKind.SearchMessages,
                                QueryHash,
                                out var decoded) &&
                            ((MessageCursorPayload)decoded).TimestampUtcTicks ==
                                TimestampUtcTicks + index;
                    }))
                    .ToArray();

                var results = await Task.WhenAll(tasks);
                Assert.That(results, Is.All.True);
            }
        }

        [Test]
        public void CursorPrimitiveAndDerivedKeyFactoryAreNotPublicSurface()
        {
            var factory = typeof(BearerToken).GetMethod(
                "CreateCursorCodec",
                BindingFlags.Instance | BindingFlags.NonPublic);

            Assert.That(typeof(HmacCursorCodec).IsNotPublic, Is.True);
            Assert.That(typeof(HmacCursorKind).IsNotPublic, Is.True);
            Assert.That(typeof(HmacCursorPayload).IsNotPublic, Is.True);
            Assert.That(factory, Is.Not.Null);
            Assert.That(factory!.IsAssembly, Is.True);
            Assert.That(
                typeof(HmacCursorCodec).GetProperties(BindingFlags.Instance | BindingFlags.Public)
                    .Any(property => property.PropertyType == typeof(byte[])),
                Is.False);
        }

        private static HmacCursorCodec CreateCodec(string tokenText = TokenText)
        {
            Assert.That(BearerToken.TryCreate(tokenText, out var token), Is.True);
            using (token)
            {
                return token.CreateCursorCodec();
            }
        }

        private static IEnumerable<TestCaseData> GetInvalidTypedPayloadCases()
        {
            yield return new TestCaseData(
                (int)HmacCursorKind.ListMailboxes,
                "{\"v\":1,\"kind\":\"listMailboxes\",\"queryHash\":\"" + QueryHash +
                "\",\"displayName\":\"" +
                new string('d', HmacCursorCodec.MaximumDisplayNameLength + 1) +
                "\",\"storeId\":\"STORE-ID\"}");
            yield return new TestCaseData(
                (int)HmacCursorKind.ListFolders,
                "{\"v\":1,\"kind\":\"listFolders\",\"queryHash\":\"" + QueryHash +
                "\",\"displayName\":\"Inbox\",\"storeId\":\"STORE-ID\"," +
                "\"entryId\":\"\"}");
            yield return new TestCaseData(
                (int)HmacCursorKind.ListMessages,
                "{\"v\":1,\"kind\":\"listMessages\",\"queryHash\":\"" + QueryHash +
                "\",\"timestampUtcTicks\":-1,\"storeId\":\"STORE-ID\"," +
                "\"entryId\":\"ENTRY-ID\",\"itemClass\":\"IPM.Note\"}");
            yield return new TestCaseData(
                (int)HmacCursorKind.SearchMessages,
                "{\"v\":1,\"kind\":\"searchMessages\",\"queryHash\":\"" + QueryHash +
                "\",\"timestampUtcTicks\":" + TimestampUtcTicks +
                ",\"storeId\":\"STORE-ID\",\"entryId\":\"ENTRY-ID\"," +
                "\"itemClass\":\"" +
                new string('i', HmacCursorCodec.MaximumItemClassLength + 1) + "\"}");
            yield return new TestCaseData(
                (int)HmacCursorKind.GetConversation,
                "{\"v\":1,\"kind\":\"getConversation\",\"queryHash\":\"" + QueryHash +
                "\",\"timestampUtcTicks\":" + TimestampUtcTicks +
                ",\"storeId\":\"STORE ID\",\"entryId\":\"ENTRY-ID\"," +
                "\"itemClass\":\"IPM.Note\"}");
            yield return new TestCaseData(
                (int)HmacCursorKind.ListAttachments,
                "{\"v\":1,\"kind\":\"listAttachments\",\"queryHash\":\"" + QueryHash +
                "\",\"attachmentIndex\":0,\"metadataFingerprint\":\"" +
                Fingerprint + "\"}");
            yield return new TestCaseData(
                (int)HmacCursorKind.ListAttachments,
                "{\"v\":1,\"kind\":\"listAttachments\",\"queryHash\":\"" + QueryHash +
                "\",\"attachmentIndex\":1,\"metadataFingerprint\":\"" +
                Fingerprint.ToUpperInvariant() + "\"}");
        }

        private static HmacCursorPayload CreatePayload(HmacCursorKind kind)
        {
            switch (kind)
            {
                case HmacCursorKind.ListMailboxes:
                    return new MailboxCursorPayload(QueryHash, "Mailbox", "STORE-ID");
                case HmacCursorKind.ListFolders:
                    return new FolderCursorPayload(
                        QueryHash,
                        "Inbox",
                        "STORE-ID",
                        "FOLDER-ID");
                case HmacCursorKind.ListMessages:
                case HmacCursorKind.SearchMessages:
                case HmacCursorKind.GetConversation:
                    return new MessageCursorPayload(
                        kind,
                        QueryHash,
                        TimestampUtcTicks,
                        "STORE-ID",
                        "ITEM-ID",
                        "IPM.Note");
                case HmacCursorKind.ListAttachments:
                    return new AttachmentCursorPayload(QueryHash, 3, Fingerprint);
                default:
                    throw new ArgumentOutOfRangeException(nameof(kind));
            }
        }

        private static void AssertPayloadEqual(
            HmacCursorPayload expected,
            HmacCursorPayload actual)
        {
            Assert.That(actual.Kind, Is.EqualTo(expected.Kind));
            Assert.That(actual.QueryHash, Is.EqualTo(expected.QueryHash));
            switch (expected)
            {
                case MailboxCursorPayload expectedMailbox:
                    var actualMailbox = (MailboxCursorPayload)actual;
                    Assert.That(actualMailbox.DisplayName, Is.EqualTo(expectedMailbox.DisplayName));
                    Assert.That(actualMailbox.StoreId, Is.EqualTo(expectedMailbox.StoreId));
                    break;
                case FolderCursorPayload expectedFolder:
                    var actualFolder = (FolderCursorPayload)actual;
                    Assert.That(actualFolder.DisplayName, Is.EqualTo(expectedFolder.DisplayName));
                    Assert.That(actualFolder.StoreId, Is.EqualTo(expectedFolder.StoreId));
                    Assert.That(actualFolder.EntryId, Is.EqualTo(expectedFolder.EntryId));
                    break;
                case MessageCursorPayload expectedMessage:
                    var actualMessage = (MessageCursorPayload)actual;
                    Assert.That(
                        actualMessage.TimestampUtcTicks,
                        Is.EqualTo(expectedMessage.TimestampUtcTicks));
                    Assert.That(actualMessage.StoreId, Is.EqualTo(expectedMessage.StoreId));
                    Assert.That(actualMessage.EntryId, Is.EqualTo(expectedMessage.EntryId));
                    Assert.That(actualMessage.ItemClass, Is.EqualTo(expectedMessage.ItemClass));
                    break;
                case AttachmentCursorPayload expectedAttachment:
                    var actualAttachment = (AttachmentCursorPayload)actual;
                    Assert.That(
                        actualAttachment.AttachmentIndex,
                        Is.EqualTo(expectedAttachment.AttachmentIndex));
                    Assert.That(
                        actualAttachment.MetadataFingerprint,
                        Is.EqualTo(expectedAttachment.MetadataFingerprint));
                    break;
                default:
                    Assert.Fail("Unexpected cursor payload type.");
                    break;
            }
        }

        private static string DecodePayloadJson(string cursor)
        {
            var segments = cursor.Split('.');
            Assert.That(segments.Length, Is.EqualTo(3));
            var bytes = DecodeBase64Url(segments[1]);
            try
            {
                return Encoding.UTF8.GetString(bytes);
            }
            finally
            {
                Array.Clear(bytes, 0, bytes.Length);
            }
        }

        private static string SignPayloadJson(string json, string tokenText)
        {
            var payloadBytes = Encoding.UTF8.GetBytes(json);
            try
            {
                return SignPayloadSegment(EncodeBase64Url(payloadBytes), tokenText);
            }
            finally
            {
                Array.Clear(payloadBytes, 0, payloadBytes.Length);
            }
        }

        private static string SignPayloadSegment(string payloadSegment, string tokenText)
        {
            var tokenBytes = DecodeBase64Url(tokenText);
            var domainBytes = Encoding.UTF8.GetBytes(KeyDomain);
            byte[] derivedKey;
            using (var derivation = new HMACSHA256(tokenBytes))
            {
                derivedKey = derivation.ComputeHash(domainBytes);
            }

            var macInput = Encoding.ASCII.GetBytes("v1." + payloadSegment);
            byte[] tag;
            using (var hmac = new HMACSHA256(derivedKey))
            {
                tag = hmac.ComputeHash(macInput);
            }

            try
            {
                return "v1." + payloadSegment + "." + EncodeBase64Url(tag);
            }
            finally
            {
                Array.Clear(tokenBytes, 0, tokenBytes.Length);
                Array.Clear(domainBytes, 0, domainBytes.Length);
                Array.Clear(derivedKey, 0, derivedKey.Length);
                Array.Clear(macInput, 0, macInput.Length);
                Array.Clear(tag, 0, tag.Length);
            }
        }

        private static Tuple<int, int> FindSeparators(string cursor)
        {
            var first = cursor.IndexOf('.');
            var second = cursor.IndexOf('.', first + 1);
            Assert.That(first, Is.EqualTo(2));
            Assert.That(second, Is.GreaterThan(first));
            return Tuple.Create(first, second);
        }

        private static string ReplaceCharacter(string value, int index)
        {
            var replacement = value[index] == 'A' ? 'B' : 'A';
            return value.Substring(0, index) + replacement + value.Substring(index + 1);
        }

        private static string FindNonCanonicalEquivalent(string canonical)
        {
            const string alphabet =
                "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_";
            var expected = DecodeBase64Url(canonical);
            try
            {
                for (var index = 0; index < alphabet.Length; index++)
                {
                    var candidate = canonical.Substring(0, canonical.Length - 1) + alphabet[index];
                    if (string.Equals(candidate, canonical, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    try
                    {
                        var decoded = DecodeBase64Url(candidate);
                        try
                        {
                            if (decoded.SequenceEqual(expected))
                            {
                                return candidate;
                            }
                        }
                        finally
                        {
                            Array.Clear(decoded, 0, decoded.Length);
                        }
                    }
                    catch (FormatException)
                    {
                    }
                }
            }
            finally
            {
                Array.Clear(expected, 0, expected.Length);
            }

            throw new InvalidOperationException("No non-canonical base64url equivalent was found.");
        }

        private static byte[] DecodeBase64Url(string value)
        {
            var paddingLength = (4 - value.Length % 4) % 4;
            return Convert.FromBase64String(
                value.Replace('-', '+').Replace('_', '/') + new string('=', paddingLength));
        }

        private static string EncodeBase64Url(byte[] value)
        {
            return Convert.ToBase64String(value)
                .TrimEnd('=')
                .Replace('+', '-')
                .Replace('/', '_');
        }
    }
}
