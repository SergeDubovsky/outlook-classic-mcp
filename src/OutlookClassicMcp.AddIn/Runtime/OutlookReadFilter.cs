using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using OutlookClassicMcp.Core.Outlook;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookClassicMcp.AddIn.Runtime
{
    internal static class OutlookReadFilter
    {
        private const string MessageClassSchema = "urn:schemas:httpmail:messageclass";
        private const string ReceivedSchema = "urn:schemas:httpmail:datereceived";
        private const string SentSchema = "urn:schemas:httpmail:date";
        private const string ModifiedSchema = "DAV:getlastmodified";
        private const string ReadSchema = "urn:schemas:httpmail:read";
        private const string SubjectSchema = "urn:schemas:httpmail:subject";
        private const string BodySchema = "urn:schemas:httpmail:textdescription";
        private const string SenderAddressSchema = "urn:schemas:httpmail:fromemail";
        private const string SenderNameSchema = "urn:schemas:httpmail:fromname";
        private const string ToSchema = "urn:schemas:httpmail:to";
        private const string CcSchema = "urn:schemas:httpmail:cc";
        private const string BccSchema = "urn:schemas:httpmail:bcc";
        private const string CategoriesSchema =
            "urn:schemas-microsoft-com:office:office#Keywords";
        private const string HasAttachmentsSchema =
            "https://schemas.microsoft.com/mapi/proptag/0x0E1B000B";

        public static string BuildListRestriction(
            OutlookMessageTimestampKind timestampKind,
            OutlookMessageKeysetAnchor? anchor)
        {
            var clauses = BaseClauses();
            if (anchor != null)
            {
                clauses.Add(
                    QuoteSchema(GetTimestampSchema(timestampKind)) +
                    " <= '" + OutlookReadDate.FormatUtc(
                        OutlookReadDate.CeilingUtcMinute(anchor.EffectiveTimestampUtc),
                        CultureInfo.CurrentCulture) + "'");
            }

            return Combine(clauses);
        }

        public static string BuildSearchRestriction(
            OutlookMessageSearchFilter filter,
            OutlookMessageKeysetAnchor? anchor,
            OutlookMessageTimestampKind timestampKind)
        {
            var clauses = BaseClauses();
            if (filter.Sender != null)
            {
                clauses.Add(
                    "(" + Contains(SenderAddressSchema, filter.Sender) +
                    " OR " + Contains(SenderNameSchema, filter.Sender) + ")");
            }

            if (filter.Recipient != null)
            {
                clauses.Add(
                    "(" + Contains(ToSchema, filter.Recipient) +
                    " OR " + Contains(CcSchema, filter.Recipient) +
                    " OR " + Contains(BccSchema, filter.Recipient) + ")");
            }

            if (filter.Subject != null)
            {
                clauses.Add(Contains(SubjectSchema, filter.Subject));
            }

            if (filter.Text != null)
            {
                clauses.Add(Contains(BodySchema, filter.Text));
            }

            if (filter.ReceivedFromUtc.HasValue)
            {
                clauses.Add(
                    QuoteSchema(ReceivedSchema) +
                    " >= '" + OutlookReadDate.FormatUtc(
                        OutlookReadDate.FloorUtcMinute(filter.ReceivedFromUtc.Value),
                        CultureInfo.CurrentCulture) + "'");
            }

            if (filter.ReceivedToUtc.HasValue)
            {
                clauses.Add(
                    QuoteSchema(ReceivedSchema) +
                    " <= '" + OutlookReadDate.FormatUtc(
                        OutlookReadDate.CeilingUtcMinute(filter.ReceivedToUtc.Value),
                        CultureInfo.CurrentCulture) + "'");
            }

            if (filter.IsUnread.HasValue)
            {
                clauses.Add(
                    QuoteSchema(ReadSchema) +
                    (filter.IsUnread.Value ? " = 0" : " = 1"));
            }

            if (filter.Category != null)
            {
                clauses.Add(
                    QuoteSchema(CategoriesSchema) +
                    " = '" + EscapeTextLiteral(filter.Category) + "'");
            }

            if (filter.HasAttachments.HasValue)
            {
                clauses.Add(
                    QuoteSchema(HasAttachmentsSchema) +
                    (filter.HasAttachments.Value ? " = 1" : " = 0"));
            }

            if (anchor != null)
            {
                clauses.Add(
                    QuoteSchema(GetTimestampSchema(timestampKind)) +
                    " <= '" + OutlookReadDate.FormatUtc(
                        OutlookReadDate.CeilingUtcMinute(anchor.EffectiveTimestampUtc),
                        CultureInfo.CurrentCulture) + "'");
            }

            return Combine(clauses);
        }

        public static bool Matches(Outlook.MailItem mail, OutlookMessageSearchFilter filter)
        {
            if (filter.Subject != null &&
                !ContainsBoundedOrdinalIgnoreCase(
                    mail.Subject,
                    filter.Subject,
                    OutlookMessageSummary.MaximumSubjectLength))
            {
                return false;
            }

            if (filter.Sender != null &&
                !ReadGuarded(
                    () => ContainsBoundedOrdinalIgnoreCase(
                            mail.SenderName,
                            filter.Sender,
                            OutlookMessageSummary.MaximumSenderLength) ||
                        ContainsBoundedOrdinalIgnoreCase(
                            mail.SenderEmailAddress,
                            filter.Sender,
                            OutlookMessageSummary.MaximumSenderLength)))
            {
                return false;
            }

            var received = OutlookReadProjection.ToOptionalUtc(mail.ReceivedTime);
            if (filter.ReceivedFromUtc.HasValue &&
                (!received.HasValue || received.Value < filter.ReceivedFromUtc.Value))
            {
                return false;
            }

            if (filter.ReceivedToUtc.HasValue &&
                (!received.HasValue || received.Value > filter.ReceivedToUtc.Value))
            {
                return false;
            }

            if (filter.IsUnread.HasValue && mail.UnRead != filter.IsUnread.Value)
            {
                return false;
            }

            if (filter.HasAttachments.HasValue)
            {
                Outlook.Attachments? attachments = null;
                try
                {
                    attachments = mail.Attachments;
                    if (attachments != null)
                    {
                        OutlookComMetrics.RecordAcquired();
                    }

                    var hasAttachments = attachments != null && attachments.Count > 0;
                    if (hasAttachments != filter.HasAttachments.Value)
                    {
                        return false;
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
            }

            return true;
        }

        public static string SortProperty(OutlookMessageTimestampKind timestampKind)
        {
            switch (timestampKind)
            {
                case OutlookMessageTimestampKind.Received:
                case OutlookMessageTimestampKind.Automatic:
                    return "[ReceivedTime]";
                case OutlookMessageTimestampKind.Sent:
                    return "[SentOn]";
                case OutlookMessageTimestampKind.Modified:
                    return "[LastModificationTime]";
                default:
                    throw new ArgumentOutOfRangeException(nameof(timestampKind));
            }
        }

        private static List<string> BaseClauses()
        {
            return new List<string>
            {
                QuoteSchema(MessageClassSchema) + " LIKE 'IPM.Note%'",
            };
        }

        private static string Combine(IEnumerable<string> clauses)
        {
            return "@SQL=" + string.Join(" AND ", clauses);
        }

        private static string Contains(string schema, string value)
        {
            return QuoteSchema(schema) + " LIKE '%" + EscapeLikeLiteral(value) + "%'";
        }

        private static string QuoteSchema(string schema)
        {
            return "\"" + schema + "\"";
        }

        private static string EscapeLikeLiteral(string value)
        {
            return EscapeTextLiteral(value)
                .Replace("[", "[[]")
                .Replace("%", "[%]")
                .Replace("_", "[_]");
        }

        private static string EscapeTextLiteral(string value)
        {
            return value.Replace("'", "''");
        }

        private static string GetTimestampSchema(OutlookMessageTimestampKind timestampKind)
        {
            switch (timestampKind)
            {
                case OutlookMessageTimestampKind.Received:
                case OutlookMessageTimestampKind.Automatic:
                    return ReceivedSchema;
                case OutlookMessageTimestampKind.Sent:
                    return SentSchema;
                case OutlookMessageTimestampKind.Modified:
                    return ModifiedSchema;
                default:
                    throw new ArgumentOutOfRangeException(nameof(timestampKind));
            }
        }

        private static bool ContainsBoundedOrdinalIgnoreCase(
            string? value,
            string expected,
            int maximumCharacters)
        {
            return value != null &&
                value.IndexOf(
                    expected,
                    0,
                    Math.Min(value.Length, maximumCharacters),
                    StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static T ReadGuarded<T>(Func<T> reader)
        {
            try
            {
                return reader();
            }
            catch (COMException exception)
                when (OutlookErrorMapper.IsObjectModelGuardDenied(exception))
            {
                throw new OutlookGatewayException(OutlookGatewayFailure.ObjectModelGuard);
            }
        }
    }
}
