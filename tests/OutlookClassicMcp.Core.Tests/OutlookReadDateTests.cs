using System;
using System.Globalization;
using NUnit.Framework;
using OutlookClassicMcp.AddIn.Runtime;

namespace OutlookClassicMcp.Core.Tests
{
    [TestFixture]
    public sealed class OutlookReadDateTests
    {
        private static readonly CultureInfo EnUs = CultureInfo.GetCultureInfo("en-US");

        [Test]
        public void FormatsTheUtcCalendarValueWithoutConvertingToTheHostTimeZone()
        {
            var utc = new DateTime(2026, 1, 15, 1, 15, 0, DateTimeKind.Utc);

            Assert.That(
                OutlookReadDate.FormatUtc(utc, EnUs),
                Is.EqualTo("1/15/2026 1:15 AM"));
        }

        [TestCase(2026, 3, 8, 7, 30, 45, "3/8/2026 7:31 AM")]
        [TestCase(2026, 11, 1, 5, 30, 45, "11/1/2026 5:31 AM")]
        public void CeilingRemainsUtcAcrossDstTransitionDates(
            int year,
            int month,
            int day,
            int hour,
            int minute,
            int second,
            string expected)
        {
            var utc = new DateTime(
                year,
                month,
                day,
                hour,
                minute,
                second,
                DateTimeKind.Utc);

            var ceiling = OutlookReadDate.CeilingUtcMinute(utc);

            Assert.That(ceiling.Kind, Is.EqualTo(DateTimeKind.Utc));
            Assert.That(OutlookReadDate.FormatUtc(ceiling, EnUs), Is.EqualTo(expected));
        }

        [Test]
        public void FloorsAtTheUtcMinuteBoundary()
        {
            var utc = new DateTime(2026, 7, 19, 22, 14, 59, DateTimeKind.Utc);

            Assert.That(
                OutlookReadDate.FloorUtcMinute(utc),
                Is.EqualTo(new DateTime(2026, 7, 19, 22, 14, 0, DateTimeKind.Utc)));
        }

        [Test]
        public void RejectsNonUtcValues()
        {
            var local = new DateTime(2026, 7, 19, 22, 14, 0, DateTimeKind.Local);
            Action format = () => OutlookReadDate.FormatUtc(local, EnUs);

            Assert.That(
                format,
                Throws.TypeOf<ArgumentException>());
        }

        [Test]
        public void CeilingDoesNotOverflowAtTheMaximumUtcMinute()
        {
            var maximum = DateTime.SpecifyKind(DateTime.MaxValue, DateTimeKind.Utc);

            Assert.That(
                OutlookReadDate.CeilingUtcMinute(maximum),
                Is.EqualTo(new DateTime(
                    9999,
                    12,
                    31,
                    23,
                    59,
                    0,
                    DateTimeKind.Utc)));
        }
    }
}
