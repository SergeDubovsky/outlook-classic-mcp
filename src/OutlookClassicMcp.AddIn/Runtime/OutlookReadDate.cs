using System;
using System.Globalization;

namespace OutlookClassicMcp.AddIn.Runtime
{
    internal static class OutlookReadDate
    {
        public static string FormatUtc(DateTime value, CultureInfo culture)
        {
            RequireUtc(value, nameof(value));
            culture = culture ?? throw new ArgumentNullException(nameof(culture));

            return value.ToString("g", culture).Replace("'", "''");
        }

        public static DateTime FloorUtcMinute(DateTime value)
        {
            RequireUtc(value, nameof(value));
            return new DateTime(
                value.Year,
                value.Month,
                value.Day,
                value.Hour,
                value.Minute,
                0,
                DateTimeKind.Utc);
        }

        public static DateTime CeilingUtcMinute(DateTime value)
        {
            var floor = FloorUtcMinute(value);
            if (floor == value || floor >= DateTime.MaxValue.AddMinutes(-1))
            {
                return floor;
            }

            return floor.AddMinutes(1);
        }

        private static void RequireUtc(DateTime value, string parameterName)
        {
            if (value.Kind != DateTimeKind.Utc)
            {
                throw new ArgumentException("The value must be UTC.", parameterName);
            }
        }
    }
}
