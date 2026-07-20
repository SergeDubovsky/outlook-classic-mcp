using System;
using System.Collections.Generic;

namespace OutlookClassicMcp.Core.Outlook
{
    internal static class OutlookContractValidation
    {
        public static string RequireBoundedText(string value, int maximumLength, string parameterName)
        {
            if (value == null)
            {
                throw new ArgumentNullException(parameterName);
            }

            if (string.IsNullOrWhiteSpace(value))
            {
                throw new ArgumentException("The value must not be empty or whitespace.", parameterName);
            }

            if (value.Length > maximumLength)
            {
                throw new ArgumentOutOfRangeException(
                    parameterName,
                    "The value exceeds the maximum permitted length.");
            }

            for (var index = 0; index < value.Length; index++)
            {
                if (char.IsControl(value[index]))
                {
                    throw new ArgumentException("The value must not contain control characters.", parameterName);
                }
            }

            return value;
        }

        public static string RequireOpaqueIdentifier(
            string value,
            int maximumLength,
            string parameterName)
        {
            var identifier = RequireBoundedText(value, maximumLength, parameterName);
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

        public static string RequireBoundedContent(
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
                throw new ArgumentOutOfRangeException(
                    parameterName,
                    "The value exceeds the maximum permitted length.");
            }

            return value;
        }

        public static string RequireBoundedDisplayText(
            string value,
            int maximumLength,
            string parameterName)
        {
            var text = RequireBoundedContent(value, maximumLength, parameterName);
            for (var index = 0; index < text.Length; index++)
            {
                if (char.IsControl(text[index]))
                {
                    throw new ArgumentException(
                        "The value must not contain control characters.",
                        parameterName);
                }
            }

            return text;
        }

        public static string RequireSha256Fingerprint(string value, string parameterName)
        {
            if (value == null)
            {
                throw new ArgumentNullException(parameterName);
            }

            if (value.Length != 64)
            {
                throw new ArgumentException(
                    "The fingerprint must be 64 lowercase hexadecimal characters.",
                    parameterName);
            }

            for (var index = 0; index < value.Length; index++)
            {
                var character = value[index];
                if (!((character >= '0' && character <= '9') ||
                    (character >= 'a' && character <= 'f')))
                {
                    throw new ArgumentException(
                        "The fingerprint must be 64 lowercase hexadecimal characters.",
                        parameterName);
                }
            }

            return value;
        }

        public static string? OptionalBoundedText(
            string? value,
            int maximumLength,
            string parameterName)
        {
            return value == null
                ? null
                : RequireBoundedText(value, maximumLength, parameterName);
        }

        public static DateTime RequireUtc(DateTime value, string parameterName)
        {
            if (value.Kind != DateTimeKind.Utc)
            {
                throw new ArgumentException("The timestamp must be UTC.", parameterName);
            }

            return value;
        }

        public static DateTime? OptionalUtc(DateTime? value, string parameterName)
        {
            return value.HasValue ? RequireUtc(value.Value, parameterName) : (DateTime?)null;
        }

        public static void RequireDefinedEnum<TEnum>(TEnum value, string parameterName)
            where TEnum : struct
        {
            if (!Enum.IsDefined(typeof(TEnum), value))
            {
                throw new ArgumentOutOfRangeException(parameterName);
            }
        }

        public static IReadOnlyList<T> BoundedCopy<T>(
            IEnumerable<T> values,
            int maximumCount,
            string parameterName,
            bool rejectNullElements)
        {
            if (values == null)
            {
                throw new ArgumentNullException(parameterName);
            }

            var copy = new List<T>();
            foreach (var value in values)
            {
                if (copy.Count == maximumCount)
                {
                    throw new ArgumentOutOfRangeException(
                        parameterName,
                        "The collection exceeds the maximum permitted count.");
                }

                if (rejectNullElements && ReferenceEquals(value, null))
                {
                    throw new ArgumentException("The collection must not contain null values.", parameterName);
                }

                copy.Add(value);
            }

            return copy.AsReadOnly();
        }
    }
}
