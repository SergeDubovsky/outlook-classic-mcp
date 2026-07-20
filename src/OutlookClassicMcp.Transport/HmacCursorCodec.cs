using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace OutlookClassicMcp.Transport
{
    internal enum HmacCursorKind
    {
        ListMailboxes = 0,
        ListFolders = 1,
        ListMessages = 2,
        SearchMessages = 3,
        GetConversation = 4,
        ListAttachments = 5,
    }

    internal abstract class HmacCursorPayload
    {
        internal HmacCursorPayload(HmacCursorKind kind, string queryHash)
        {
            HmacCursorValidation.RequireDefinedKind(kind, nameof(kind));
            Kind = kind;
            QueryHash = HmacCursorValidation.RequireSha256Hex(queryHash, nameof(queryHash));
        }

        public HmacCursorKind Kind { get; }

        public string QueryHash { get; }
    }

    internal sealed class MailboxCursorPayload : HmacCursorPayload
    {
        internal MailboxCursorPayload(string queryHash, string displayName, string storeId)
            : base(HmacCursorKind.ListMailboxes, queryHash)
        {
            DisplayName = HmacCursorValidation.RequireDisplayText(
                displayName,
                HmacCursorCodec.MaximumDisplayNameLength,
                nameof(displayName));
            StoreId = HmacCursorValidation.RequireOpaqueIdentifier(storeId, nameof(storeId));
        }

        public string DisplayName { get; }

        public string StoreId { get; }
    }

    internal sealed class FolderCursorPayload : HmacCursorPayload
    {
        internal FolderCursorPayload(
            string queryHash,
            string displayName,
            string storeId,
            string entryId)
            : base(HmacCursorKind.ListFolders, queryHash)
        {
            DisplayName = HmacCursorValidation.RequireDisplayText(
                displayName,
                HmacCursorCodec.MaximumDisplayNameLength,
                nameof(displayName));
            StoreId = HmacCursorValidation.RequireOpaqueIdentifier(storeId, nameof(storeId));
            EntryId = HmacCursorValidation.RequireOpaqueIdentifier(entryId, nameof(entryId));
        }

        public string DisplayName { get; }

        public string StoreId { get; }

        public string EntryId { get; }
    }

    internal sealed class MessageCursorPayload : HmacCursorPayload
    {
        internal MessageCursorPayload(
            HmacCursorKind kind,
            string queryHash,
            long timestampUtcTicks,
            string storeId,
            string entryId,
            string itemClass)
            : base(RequireMessageKind(kind), queryHash)
        {
            if (timestampUtcTicks < DateTime.MinValue.Ticks ||
                timestampUtcTicks > DateTime.MaxValue.Ticks)
            {
                throw new ArgumentOutOfRangeException(nameof(timestampUtcTicks));
            }

            TimestampUtcTicks = timestampUtcTicks;
            StoreId = HmacCursorValidation.RequireOpaqueIdentifier(storeId, nameof(storeId));
            EntryId = HmacCursorValidation.RequireOpaqueIdentifier(entryId, nameof(entryId));
            ItemClass = HmacCursorValidation.RequireBoundedText(
                itemClass,
                HmacCursorCodec.MaximumItemClassLength,
                nameof(itemClass));
        }

        public long TimestampUtcTicks { get; }

        public string StoreId { get; }

        public string EntryId { get; }

        public string ItemClass { get; }

        private static HmacCursorKind RequireMessageKind(HmacCursorKind kind)
        {
            switch (kind)
            {
                case HmacCursorKind.ListMessages:
                case HmacCursorKind.SearchMessages:
                case HmacCursorKind.GetConversation:
                    return kind;
                default:
                    throw new ArgumentOutOfRangeException(nameof(kind));
            }
        }
    }

    internal sealed class AttachmentCursorPayload : HmacCursorPayload
    {
        internal AttachmentCursorPayload(
            string queryHash,
            int attachmentIndex,
            string metadataFingerprint)
            : base(HmacCursorKind.ListAttachments, queryHash)
        {
            if (attachmentIndex < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(attachmentIndex));
            }

            AttachmentIndex = attachmentIndex;
            MetadataFingerprint = HmacCursorValidation.RequireSha256Hex(
                metadataFingerprint,
                nameof(metadataFingerprint));
        }

        public int AttachmentIndex { get; }

        public string MetadataFingerprint { get; }
    }

    internal sealed class HmacCursorCodec : IDisposable
    {
        // The Core locator contract permits two 4,096-character identifiers.
        // Utf8JsonWriter can escape each UTF-16 code unit as six ASCII bytes,
        // so the authenticated base64url envelope needs a larger closed bound.
        public const int MaximumCursorLength = 96 * 1024;
        public const int MaximumDisplayNameLength = 256;
        public const int MaximumOpaqueIdentifierLength = 4096;
        public const int MaximumItemClassLength = 256;
        public const int Sha256HexLength = 64;

        private const string FormatVersion = "v1";
        private const int PayloadVersion = 1;
        private const int KeyLength = 32;
        private const int TagLength = 32;
        private const int EncodedTagLength = 43;
        private readonly object _gate = new object();
        private byte[]? _key;

        internal HmacCursorCodec(byte[] key)
        {
            if (key == null)
            {
                throw new ArgumentNullException(nameof(key));
            }

            if (key.Length != KeyLength)
            {
                throw new ArgumentException("The cursor key must contain 32 bytes.", nameof(key));
            }

            _key = new byte[KeyLength];
            Buffer.BlockCopy(key, 0, _key, 0, KeyLength);
        }

        internal string Encode(HmacCursorPayload payload)
        {
            if (payload == null)
            {
                throw new ArgumentNullException(nameof(payload));
            }

            var key = CopyKey();
            byte[] payloadBytes = Array.Empty<byte>();
            byte[] macInput = Array.Empty<byte>();
            byte[] tag = Array.Empty<byte>();
            try
            {
                payloadBytes = SerializePayload(payload);
                var payloadSegment = EncodeBase64Url(payloadBytes);
                var authenticatedPrefix = FormatVersion + "." + payloadSegment;
                macInput = Encoding.ASCII.GetBytes(authenticatedPrefix);
                tag = ComputeTag(key, macInput);
                var cursor = authenticatedPrefix + "." + EncodeBase64Url(tag);
                if (cursor.Length > MaximumCursorLength)
                {
                    throw new ArgumentOutOfRangeException(
                        nameof(payload),
                        "The encoded cursor exceeds the maximum permitted length.");
                }

                return cursor;
            }
            finally
            {
                Clear(key);
                Clear(payloadBytes);
                Clear(macInput);
                Clear(tag);
            }
        }

        internal bool TryDecode(
            string? cursor,
            HmacCursorKind expectedKind,
            string expectedQueryHash,
            out HmacCursorPayload payload)
        {
            payload = null!;
            var key = CopyKey();
            try
            {
                if (!HmacCursorValidation.IsDefinedKind(expectedKind) ||
                    !HmacCursorValidation.IsSha256Hex(expectedQueryHash) ||
                    cursor == null ||
                    cursor.Length > MaximumCursorLength ||
                    !TrySplit(cursor, out var payloadSegment, out var tagSegment) ||
                    !IsBase64UrlText(payloadSegment) ||
                    tagSegment.Length != EncodedTagLength)
                {
                    return false;
                }

                if (!TryAuthenticate(key, payloadSegment, tagSegment))
                {
                    return false;
                }

                if (!TryDecodeBase64Url(payloadSegment, out var payloadBytes))
                {
                    return false;
                }

                try
                {
                    if (!TryParsePayload(
                        payloadBytes,
                        expectedKind,
                        expectedQueryHash,
                        out var parsedPayload))
                    {
                        return false;
                    }

                    byte[] canonicalBytes;
                    try
                    {
                        canonicalBytes = SerializePayload(parsedPayload);
                    }
                    catch (ArgumentException)
                    {
                        return false;
                    }
                    catch (InvalidOperationException)
                    {
                        return false;
                    }

                    try
                    {
                        if (!BytesEqual(payloadBytes, canonicalBytes))
                        {
                            return false;
                        }
                    }
                    finally
                    {
                        Clear(canonicalBytes);
                    }

                    payload = parsedPayload;
                    return true;
                }
                finally
                {
                    Clear(payloadBytes);
                }
            }
            finally
            {
                Clear(key);
            }
        }

        public void Dispose()
        {
            byte[]? key;
            lock (_gate)
            {
                key = _key;
                _key = null;
            }

            if (key != null)
            {
                Clear(key);
            }
        }

        private byte[] CopyKey()
        {
            lock (_gate)
            {
                var key = _key ?? throw new ObjectDisposedException(nameof(HmacCursorCodec));
                var copy = new byte[key.Length];
                Buffer.BlockCopy(key, 0, copy, 0, key.Length);
                return copy;
            }
        }

        private static bool TryAuthenticate(
            byte[] key,
            string payloadSegment,
            string tagSegment)
        {
            if (!TryDecodeBase64Url(tagSegment, out var candidateTag) ||
                candidateTag.Length != TagLength)
            {
                Clear(candidateTag);
                return false;
            }

            byte[] macInput = Array.Empty<byte>();
            byte[] expectedTag = Array.Empty<byte>();
            try
            {
                macInput = Encoding.ASCII.GetBytes(FormatVersion + "." + payloadSegment);
                expectedTag = ComputeTag(key, macInput);
                return FixedTimeEquals(expectedTag, candidateTag);
            }
            finally
            {
                Clear(candidateTag);
                Clear(macInput);
                Clear(expectedTag);
            }
        }

        private static byte[] ComputeTag(byte[] key, byte[] value)
        {
            using (var hmac = new HMACSHA256(key))
            {
                return hmac.ComputeHash(value);
            }
        }

        private static bool FixedTimeEquals(byte[] expected, byte[] candidate)
        {
            if (expected.Length != candidate.Length)
            {
                return false;
            }

            var difference = 0;
            for (var index = 0; index < expected.Length; index++)
            {
                difference |= expected[index] ^ candidate[index];
            }

            return difference == 0;
        }

        private static bool TrySplit(
            string cursor,
            out string payloadSegment,
            out string tagSegment)
        {
            payloadSegment = string.Empty;
            tagSegment = string.Empty;
            var firstSeparator = cursor.IndexOf('.');
            if (firstSeparator != FormatVersion.Length ||
                !cursor.StartsWith(FormatVersion + ".", StringComparison.Ordinal))
            {
                return false;
            }

            var secondSeparator = cursor.IndexOf('.', firstSeparator + 1);
            if (secondSeparator <= firstSeparator + 1 ||
                secondSeparator == cursor.Length - 1 ||
                cursor.IndexOf('.', secondSeparator + 1) >= 0)
            {
                return false;
            }

            payloadSegment = cursor.Substring(
                firstSeparator + 1,
                secondSeparator - firstSeparator - 1);
            tagSegment = cursor.Substring(secondSeparator + 1);
            return true;
        }

        private static byte[] SerializePayload(HmacCursorPayload payload)
        {
            using (var stream = new MemoryStream())
            {
                try
                {
                    using (var writer = new Utf8JsonWriter(stream))
                    {
                        writer.WriteStartObject();
                        writer.WriteNumber("v", PayloadVersion);
                        writer.WriteString("kind", GetKindLabel(payload.Kind));
                        writer.WriteString("queryHash", payload.QueryHash);
                        switch (payload)
                        {
                            case MailboxCursorPayload mailbox:
                                writer.WriteString("displayName", mailbox.DisplayName);
                                writer.WriteString("storeId", mailbox.StoreId);
                                break;
                            case FolderCursorPayload folder:
                                writer.WriteString("displayName", folder.DisplayName);
                                writer.WriteString("storeId", folder.StoreId);
                                writer.WriteString("entryId", folder.EntryId);
                                break;
                            case MessageCursorPayload message:
                                writer.WriteNumber("timestampUtcTicks", message.TimestampUtcTicks);
                                writer.WriteString("storeId", message.StoreId);
                                writer.WriteString("entryId", message.EntryId);
                                writer.WriteString("itemClass", message.ItemClass);
                                break;
                            case AttachmentCursorPayload attachment:
                                writer.WriteNumber("attachmentIndex", attachment.AttachmentIndex);
                                writer.WriteString(
                                    "metadataFingerprint",
                                    attachment.MetadataFingerprint);
                                break;
                            default:
                                throw new ArgumentException(
                                    "The cursor payload type is unsupported.",
                                    nameof(payload));
                        }

                        writer.WriteEndObject();
                        writer.Flush();
                    }

                    return stream.ToArray();
                }
                finally
                {
                    if (stream.TryGetBuffer(out var buffer) && buffer.Array != null)
                    {
                        Array.Clear(buffer.Array, buffer.Offset, buffer.Count);
                    }
                }
            }
        }

        private static bool TryParsePayload(
            byte[] payloadBytes,
            HmacCursorKind expectedKind,
            string expectedQueryHash,
            out HmacCursorPayload payload)
        {
            payload = null!;
            try
            {
                using (var document = JsonDocument.Parse(
                    payloadBytes,
                    new JsonDocumentOptions
                    {
                        AllowTrailingCommas = false,
                        CommentHandling = JsonCommentHandling.Disallow,
                        MaxDepth = 8,
                    }))
                {
                    var root = document.RootElement;
                    if (root.ValueKind != JsonValueKind.Object ||
                        !TryGetInt32(root, "v", out var version) ||
                        version != PayloadVersion ||
                        !TryGetString(root, "kind", out var kindLabel) ||
                        !string.Equals(
                            kindLabel,
                            GetKindLabel(expectedKind),
                            StringComparison.Ordinal) ||
                        !TryGetString(root, "queryHash", out var queryHash) ||
                        !HmacCursorValidation.IsSha256Hex(queryHash) ||
                        !string.Equals(queryHash, expectedQueryHash, StringComparison.Ordinal))
                    {
                        return false;
                    }

                    switch (expectedKind)
                    {
                        case HmacCursorKind.ListMailboxes:
                            if (!TryGetString(root, "displayName", out var mailboxName) ||
                                !TryGetString(root, "storeId", out var mailboxStoreId))
                            {
                                return false;
                            }

                            payload = new MailboxCursorPayload(
                                queryHash,
                                mailboxName,
                                mailboxStoreId);
                            return true;
                        case HmacCursorKind.ListFolders:
                            if (!TryGetString(root, "displayName", out var folderName) ||
                                !TryGetString(root, "storeId", out var folderStoreId) ||
                                !TryGetString(root, "entryId", out var folderEntryId))
                            {
                                return false;
                            }

                            payload = new FolderCursorPayload(
                                queryHash,
                                folderName,
                                folderStoreId,
                                folderEntryId);
                            return true;
                        case HmacCursorKind.ListMessages:
                        case HmacCursorKind.SearchMessages:
                        case HmacCursorKind.GetConversation:
                            if (!TryGetInt64(root, "timestampUtcTicks", out var ticks) ||
                                !TryGetString(root, "storeId", out var messageStoreId) ||
                                !TryGetString(root, "entryId", out var messageEntryId) ||
                                !TryGetString(root, "itemClass", out var itemClass))
                            {
                                return false;
                            }

                            payload = new MessageCursorPayload(
                                expectedKind,
                                queryHash,
                                ticks,
                                messageStoreId,
                                messageEntryId,
                                itemClass);
                            return true;
                        case HmacCursorKind.ListAttachments:
                            if (!TryGetInt32(root, "attachmentIndex", out var attachmentIndex) ||
                                !TryGetString(
                                    root,
                                    "metadataFingerprint",
                                    out var metadataFingerprint))
                            {
                                return false;
                            }

                            payload = new AttachmentCursorPayload(
                                queryHash,
                                attachmentIndex,
                                metadataFingerprint);
                            return true;
                        default:
                            return false;
                    }
                }
            }
            catch (JsonException)
            {
                return false;
            }
            catch (ArgumentException)
            {
                return false;
            }
            catch (InvalidOperationException)
            {
                return false;
            }
        }

        private static bool TryGetString(
            JsonElement value,
            string propertyName,
            out string result)
        {
            result = string.Empty;
            if (!value.TryGetProperty(propertyName, out var property) ||
                property.ValueKind != JsonValueKind.String)
            {
                return false;
            }

            var text = property.GetString();
            if (text == null)
            {
                return false;
            }

            result = text;
            return true;
        }

        private static bool TryGetInt32(
            JsonElement value,
            string propertyName,
            out int result)
        {
            result = 0;
            return value.TryGetProperty(propertyName, out var property) &&
                property.ValueKind == JsonValueKind.Number &&
                property.TryGetInt32(out result);
        }

        private static bool TryGetInt64(
            JsonElement value,
            string propertyName,
            out long result)
        {
            result = 0;
            return value.TryGetProperty(propertyName, out var property) &&
                property.ValueKind == JsonValueKind.Number &&
                property.TryGetInt64(out result);
        }

        private static string GetKindLabel(HmacCursorKind kind)
        {
            switch (kind)
            {
                case HmacCursorKind.ListMailboxes:
                    return "listMailboxes";
                case HmacCursorKind.ListFolders:
                    return "listFolders";
                case HmacCursorKind.ListMessages:
                    return "listMessages";
                case HmacCursorKind.SearchMessages:
                    return "searchMessages";
                case HmacCursorKind.GetConversation:
                    return "getConversation";
                case HmacCursorKind.ListAttachments:
                    return "listAttachments";
                default:
                    throw new ArgumentOutOfRangeException(nameof(kind));
            }
        }

        private static string EncodeBase64Url(byte[] value)
        {
            return Convert.ToBase64String(value)
                .TrimEnd('=')
                .Replace('+', '-')
                .Replace('/', '_');
        }

        private static bool TryDecodeBase64Url(string encoded, out byte[] value)
        {
            value = Array.Empty<byte>();
            if (!IsBase64UrlText(encoded) || encoded.Length % 4 == 1)
            {
                return false;
            }

            try
            {
                var paddingLength = (4 - encoded.Length % 4) % 4;
                var base64 = encoded.Replace('-', '+').Replace('_', '/') +
                    new string('=', paddingLength);
                var decoded = Convert.FromBase64String(base64);
                if (!string.Equals(EncodeBase64Url(decoded), encoded, StringComparison.Ordinal))
                {
                    Clear(decoded);
                    return false;
                }

                value = decoded;
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }

        private static bool IsBase64UrlText(string value)
        {
            if (value.Length == 0)
            {
                return false;
            }

            for (var index = 0; index < value.Length; index++)
            {
                var character = value[index];
                if (!((character >= 'A' && character <= 'Z') ||
                    (character >= 'a' && character <= 'z') ||
                    (character >= '0' && character <= '9') ||
                    character == '-' ||
                    character == '_'))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool BytesEqual(byte[] left, byte[] right)
        {
            if (left.Length != right.Length)
            {
                return false;
            }

            for (var index = 0; index < left.Length; index++)
            {
                if (left[index] != right[index])
                {
                    return false;
                }
            }

            return true;
        }

        private static void Clear(byte[] value)
        {
            Array.Clear(value, 0, value.Length);
        }
    }

    internal static class HmacCursorValidation
    {
        public static bool IsDefinedKind(HmacCursorKind kind)
        {
            return kind >= HmacCursorKind.ListMailboxes &&
                kind <= HmacCursorKind.ListAttachments;
        }

        public static void RequireDefinedKind(HmacCursorKind kind, string parameterName)
        {
            if (!IsDefinedKind(kind))
            {
                throw new ArgumentOutOfRangeException(parameterName);
            }
        }

        public static string RequireSha256Hex(string value, string parameterName)
        {
            if (!IsSha256Hex(value))
            {
                throw new ArgumentException(
                    "The value must be 64 lowercase hexadecimal characters.",
                    parameterName);
            }

            return value;
        }

        public static bool IsSha256Hex(string? value)
        {
            if (value == null || value.Length != HmacCursorCodec.Sha256HexLength)
            {
                return false;
            }

            for (var index = 0; index < value.Length; index++)
            {
                var character = value[index];
                if (!((character >= '0' && character <= '9') ||
                    (character >= 'a' && character <= 'f')))
                {
                    return false;
                }
            }

            return true;
        }

        public static string RequireOpaqueIdentifier(string value, string parameterName)
        {
            var identifier = RequireBoundedText(
                value,
                HmacCursorCodec.MaximumOpaqueIdentifierLength,
                parameterName);
            for (var index = 0; index < identifier.Length; index++)
            {
                if (char.IsWhiteSpace(identifier[index]))
                {
                    throw new ArgumentException(
                        "The identifier must not contain whitespace.",
                        parameterName);
                }
            }

            return identifier;
        }

        public static string RequireDisplayText(
            string value,
            int maximumLength,
            string parameterName)
        {
            if (value == null)
            {
                throw new ArgumentNullException(parameterName);
            }

            if (value.Length > maximumLength)
            {
                throw new ArgumentOutOfRangeException(parameterName);
            }

            for (var index = 0; index < value.Length; index++)
            {
                if (char.IsControl(value[index]))
                {
                    throw new ArgumentException(
                        "The value contains control characters.",
                        parameterName);
                }
            }

            return value;
        }

        public static string RequireBoundedText(
            string value,
            int maximumLength,
            string parameterName)
        {
            if (value == null)
            {
                throw new ArgumentNullException(parameterName);
            }

            if (string.IsNullOrWhiteSpace(value))
            {
                throw new ArgumentException(
                    "The value must not be empty or whitespace.",
                    parameterName);
            }

            if (value.Length > maximumLength)
            {
                throw new ArgumentOutOfRangeException(parameterName);
            }

            ValidateText(value, parameterName);
            return value;
        }

        private static void ValidateText(string value, string parameterName)
        {
            for (var index = 0; index < value.Length; index++)
            {
                var character = value[index];
                if (char.IsControl(character))
                {
                    throw new ArgumentException(
                        "The value contains control characters.",
                        parameterName);
                }
            }
        }
    }
}
