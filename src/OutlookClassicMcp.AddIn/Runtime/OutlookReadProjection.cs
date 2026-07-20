using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using OutlookClassicMcp.Core.Outlook;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookClassicMcp.AddIn.Runtime
{
    internal enum OutlookMessageTimestampKind
    {
        Received = 0,
        Sent = 1,
        Modified = 2,
        Automatic = 3,
    }

    internal static class OutlookReadProjection
    {
        private const int MapiNotSupported = unchecked((int)0x80040102);
        private static readonly DateTime MinimumUsefulDate = new DateTime(1900, 1, 1);
        private static readonly DateTime MaximumUsefulDate = new DateTime(4500, 1, 1);

        public static OutlookMessageSummary ProjectMessage(
            Outlook.MailItem mail,
            FolderRef folder,
            OutlookMessageTimestampKind timestampKind)
        {
            var entryId = RequireIdentifier(mail.EntryID, ItemRef.MaximumEntryIdLength);
            var itemClass = RequireText(mail.MessageClass, ItemRef.MaximumItemClassLength);
            var subject = BoundDisplay(mail.Subject, OutlookMessageSummary.MaximumSubjectLength);
            var senderName = ReadGuardedOptional(
                () => mail.SenderName,
                OutlookMessageSummary.MaximumSenderLength);
            var senderAddress = ReadGuardedOptional(
                () => mail.SenderEmailAddress,
                OutlookMessageSummary.MaximumSenderLength);
            var receivedUtc = ToOptionalUtc(mail.ReceivedTime);
            var sentUtc = ToOptionalUtc(mail.SentOn);
            var modifiedUtc = ToOptionalUtc(mail.LastModificationTime);
            var effectiveTimestampUtc = SelectEffectiveTimestamp(
                timestampKind,
                receivedUtc,
                sentUtc,
                modifiedUtc);

            Outlook.Attachments? attachments = null;
            var attachmentCount = 0;
            try
            {
                attachments = mail.Attachments;
                if (attachments != null)
                {
                    OutlookComMetrics.RecordAcquired();
                    attachmentCount = attachments.Count;
                }
            }
            finally
            {
                if (attachments != null)
                {
                    Marshal.ReleaseComObject(attachments);
                    OutlookComMetrics.RecordReleased();
                }
            }

            string? conversationId;
            try
            {
                conversationId = BoundOptionalIdentifier(
                    mail.ConversationID,
                    OutlookMessageSummary.MaximumConversationIdLength);
            }
            catch (COMException exception) when (!IsFatalProviderFailure(exception))
            {
                conversationId = null;
            }

            return new OutlookMessageSummary(
                new ItemRef(folder.StoreId, entryId, itemClass),
                folder,
                subject,
                senderName,
                senderAddress,
                effectiveTimestampUtc,
                receivedUtc,
                sentUtc,
                !mail.UnRead,
                attachmentCount,
                conversationId);
        }

        public static OutlookStoreType MapStoreType(Outlook.OlExchangeStoreType value)
        {
            switch (value)
            {
                case Outlook.OlExchangeStoreType.olPrimaryExchangeMailbox:
                    return OutlookStoreType.PrimaryExchangeMailbox;
                case Outlook.OlExchangeStoreType.olExchangeMailbox:
                    return OutlookStoreType.ExchangeMailbox;
                case Outlook.OlExchangeStoreType.olExchangePublicFolder:
                    return OutlookStoreType.ExchangePublicFolder;
                case Outlook.OlExchangeStoreType.olAdditionalExchangeMailbox:
                    return OutlookStoreType.AdditionalExchangeMailbox;
                case Outlook.OlExchangeStoreType.olNotExchange:
                    return OutlookStoreType.NonExchange;
                default:
                    return OutlookStoreType.Unknown;
            }
        }

        public static DateTime SelectEffectiveTimestamp(
            OutlookMessageTimestampKind timestampKind,
            DateTime? receivedUtc,
            DateTime? sentUtc,
            DateTime? modifiedUtc)
        {
            DateTime? selected;
            switch (timestampKind)
            {
                case OutlookMessageTimestampKind.Received:
                    selected = receivedUtc;
                    break;
                case OutlookMessageTimestampKind.Sent:
                    selected = sentUtc;
                    break;
                case OutlookMessageTimestampKind.Modified:
                    selected = modifiedUtc;
                    break;
                case OutlookMessageTimestampKind.Automatic:
                    selected = receivedUtc ?? sentUtc ?? modifiedUtc;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(timestampKind));
            }

            return selected ?? DateTime.SpecifyKind(DateTime.MinValue, DateTimeKind.Utc);
        }

        public static DateTime ReadOrderingTimestamp(
            Outlook.MailItem mail,
            OutlookMessageTimestampKind timestampKind)
        {
            DateTime? value;
            switch (timestampKind)
            {
                case OutlookMessageTimestampKind.Received:
                    value = ToOptionalUtc(mail.ReceivedTime);
                    break;
                case OutlookMessageTimestampKind.Sent:
                    value = ToOptionalUtc(mail.SentOn);
                    break;
                case OutlookMessageTimestampKind.Modified:
                    value = ToOptionalUtc(mail.LastModificationTime);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(timestampKind));
            }

            return value ?? DateTime.SpecifyKind(DateTime.MinValue, DateTimeKind.Utc);
        }

        public static DateTime? ToOptionalUtc(DateTime value)
        {
            if (value < MinimumUsefulDate || value >= MaximumUsefulDate)
            {
                return null;
            }

            if (value.Kind == DateTimeKind.Utc)
            {
                return value;
            }

            var local = value.Kind == DateTimeKind.Local
                ? value
                : DateTime.SpecifyKind(value, DateTimeKind.Local);
            return local.ToUniversalTime();
        }

        public static FolderRef ReadParentFolder(Outlook.MailItem mail)
        {
            object? parent = null;
            try
            {
                parent = mail.Parent;
                if (parent != null && Marshal.IsComObject(parent))
                {
                    OutlookComMetrics.RecordAcquired();
                }

                if (!(parent is Outlook.MAPIFolder folder))
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.ItemMovedOrDeleted);
                }

                return new FolderRef(
                    RequireIdentifier(folder.StoreID, MailboxRef.MaximumStoreIdLength),
                    RequireIdentifier(folder.EntryID, FolderRef.MaximumEntryIdLength));
            }
            finally
            {
                if (parent != null && Marshal.IsComObject(parent))
                {
                    Marshal.ReleaseComObject(parent);
                    OutlookComMetrics.RecordReleased();
                }
            }
        }

        public static int CompareMessages(
            OutlookMessageSummary left,
            OutlookMessageSummary right)
        {
            var timestamp = right.EffectiveTimestampUtc.CompareTo(left.EffectiveTimestampUtc);
            if (timestamp != 0)
            {
                return timestamp;
            }

            var store = string.CompareOrdinal(left.Item.StoreId, right.Item.StoreId);
            return store != 0
                ? store
                : string.CompareOrdinal(left.Item.EntryId, right.Item.EntryId);
        }

        public static int CompareMessageToAnchor(
            OutlookMessageSummary value,
            OutlookMessageKeysetAnchor anchor)
        {
            var timestamp = anchor.EffectiveTimestampUtc.CompareTo(value.EffectiveTimestampUtc);
            if (timestamp != 0)
            {
                return timestamp;
            }

            var store = string.CompareOrdinal(value.Item.StoreId, anchor.Item.StoreId);
            return store != 0
                ? store
                : string.CompareOrdinal(value.Item.EntryId, anchor.Item.EntryId);
        }

        public static int CompareMailboxes(
            OutlookMailboxSummary left,
            OutlookMailboxSummary right)
        {
            var name = string.CompareOrdinal(left.DisplayName, right.DisplayName);
            return name != 0
                ? name
                : string.CompareOrdinal(left.Mailbox.StoreId, right.Mailbox.StoreId);
        }

        public static int CompareMailboxToAnchor(
            OutlookMailboxSummary value,
            OutlookMailboxKeysetAnchor anchor)
        {
            var name = string.CompareOrdinal(value.DisplayName, anchor.DisplayName);
            return name != 0
                ? name
                : string.CompareOrdinal(value.Mailbox.StoreId, anchor.Mailbox.StoreId);
        }

        public static int CompareFolders(OutlookFolderSummary left, OutlookFolderSummary right)
        {
            var name = string.CompareOrdinal(left.DisplayName, right.DisplayName);
            if (name != 0)
            {
                return name;
            }

            var store = string.CompareOrdinal(left.Folder.StoreId, right.Folder.StoreId);
            return store != 0
                ? store
                : string.CompareOrdinal(left.Folder.EntryId, right.Folder.EntryId);
        }

        public static int CompareFolderToAnchor(
            OutlookFolderSummary value,
            OutlookFolderKeysetAnchor anchor)
        {
            var name = string.CompareOrdinal(value.DisplayName, anchor.DisplayName);
            if (name != 0)
            {
                return name;
            }

            var store = string.CompareOrdinal(value.Folder.StoreId, anchor.Folder.StoreId);
            return store != 0
                ? store
                : string.CompareOrdinal(value.Folder.EntryId, anchor.Folder.EntryId);
        }

        public static void InsertBounded<T>(
            List<T> values,
            T value,
            int maximumCount,
            Comparison<T> comparison)
        {
            var index = values.BinarySearch(value, Comparer<T>.Create(comparison));
            if (index < 0)
            {
                index = ~index;
            }

            values.Insert(index, value);
            if (values.Count > maximumCount)
            {
                values.RemoveAt(values.Count - 1);
            }
        }

        public static OutlookMessagePage BuildMessagePage(
            List<OutlookMessageSummary> candidates,
            int pageSize,
            int totalScopeCount,
            IEnumerable<OutlookScopeFailure> failures)
        {
            var failureList = new List<OutlookScopeFailure>(failures);
            candidates.Sort(CompareMessages);
            var hasMore = candidates.Count > pageSize;
            if (hasMore)
            {
                candidates.RemoveRange(pageSize, candidates.Count - pageSize);
            }

            OutlookMessageKeysetAnchor? nextAnchor = null;
            if (hasMore && failureList.Count == 0 && candidates.Count > 0)
            {
                var last = candidates[candidates.Count - 1];
                nextAnchor = new OutlookMessageKeysetAnchor(
                    last.EffectiveTimestampUtc,
                    last.Item);
            }

            return new OutlookMessagePage(
                candidates,
                nextAnchor,
                totalScopeCount,
                failureList);
        }

        public static string BoundDisplay(string? value, int maximumLength)
        {
            if (value == null)
            {
                return string.Empty;
            }

            var builder = new StringBuilder(Math.Min(value.Length, maximumLength));
            for (var index = 0; index < value.Length && builder.Length < maximumLength; index++)
            {
                if (!char.IsControl(value[index]))
                {
                    builder.Append(value[index]);
                }
            }

            return builder.ToString();
        }

        public static string? BoundOptionalText(string? value, int maximumLength)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            var bounded = BoundDisplay(value, maximumLength).Trim();
            return bounded.Length == 0 ? null : bounded;
        }

        public static string RequireText(string? value, int maximumLength)
        {
            if (string.IsNullOrWhiteSpace(value) || value!.Length > maximumLength)
            {
                throw new OutlookGatewayException(OutlookGatewayFailure.Internal);
            }

            for (var index = 0; index < value.Length; index++)
            {
                if (char.IsControl(value[index]))
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.Internal);
                }
            }

            return value;
        }

        public static string RequireIdentifier(string? value, int maximumLength)
        {
            var identifier = RequireText(value, maximumLength);
            for (var index = 0; index < identifier.Length; index++)
            {
                if (char.IsWhiteSpace(identifier[index]))
                {
                    throw new OutlookGatewayException(OutlookGatewayFailure.Internal);
                }
            }

            return identifier;
        }

        public static string ComputeAttachmentFingerprint(
            ItemRef item,
            int index,
            string name,
            string? displayName,
            long size,
            int type,
            int position,
            int blockLevel,
            string? contentType)
        {
            using (var stream = new MemoryStream())
            using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true))
            {
                writer.Write((byte)1);
                writer.Write(item.StoreId);
                writer.Write(item.EntryId);
                writer.Write(item.ItemClass);
                writer.Write(index);
                writer.Write(name);
                writer.Write(displayName ?? string.Empty);
                writer.Write(size);
                writer.Write(type);
                writer.Write(position);
                writer.Write(blockLevel);
                writer.Write(contentType ?? string.Empty);
                writer.Flush();

                using (var sha256 = SHA256.Create())
                {
                    var hash = sha256.ComputeHash(stream.ToArray());
                    var builder = new StringBuilder(hash.Length * 2);
                    foreach (var value in hash)
                    {
                        builder.Append(value.ToString("x2", CultureInfo.InvariantCulture));
                    }

                    return builder.ToString();
                }
            }
        }

        private static string? ReadGuardedOptional(
            Func<string> reader,
            int maximumLength)
        {
            try
            {
                return BoundOptionalText(reader(), maximumLength);
            }
            catch (COMException exception)
                when (exception.ErrorCode == MapiNotSupported ||
                    exception.ErrorCode == unchecked((int)0x80070005))
            {
                return null;
            }
        }

        private static string? BoundOptionalIdentifier(string? value, int maximumLength)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            if (value!.Length > maximumLength)
            {
                return null;
            }

            for (var index = 0; index < value.Length; index++)
            {
                if (char.IsWhiteSpace(value[index]))
                {
                    return null;
                }
            }

            return value;
        }

        private static bool IsFatalProviderFailure(COMException exception)
        {
            switch (exception.ErrorCode)
            {
                case unchecked((int)0x80010001):
                case unchecked((int)0x8001010A):
                case unchecked((int)0x80010108):
                case unchecked((int)0x800401FD):
                    return true;
                default:
                    return false;
            }
        }
    }
}
