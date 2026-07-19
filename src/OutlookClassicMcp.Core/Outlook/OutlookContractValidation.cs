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
