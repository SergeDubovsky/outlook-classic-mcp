#if NET48
using System;
using System.Globalization;
using System.Runtime.InteropServices;
using NUnit.Framework;
using OutlookClassicMcp.AddIn.Runtime;
using OutlookClassicMcp.Core.Outlook;

namespace OutlookClassicMcp.Core.Tests
{
    [TestFixture]
    public sealed class OutlookErrorMapperTests
    {
        [TestCase("HOST_BUSY", OutlookGatewayFailure.QueueFull)]
        [TestCase("HOST_STOPPING", OutlookGatewayFailure.Stopping)]
        [TestCase("HOST_UNAVAILABLE", OutlookGatewayFailure.StaDispatchFailed)]
        [TestCase(
            "Outlook work executed outside the captured UI STA.",
            OutlookGatewayFailure.StaDispatchFailed)]
        public void MapsDispatcherFailures(string message, OutlookGatewayFailure expected)
        {
            var mapped = OutlookErrorMapper.Map(new InvalidOperationException(message));

            Assert.That(mapped.Failure, Is.EqualTo(expected));
            Assert.That(mapped.InnerException, Is.Null);
            Assert.That(mapped.Message, Does.Not.Contain(message));
        }

        [TestCase(unchecked((int)0x80010001), OutlookGatewayFailure.ComBusy)]
        [TestCase(unchecked((int)0x8001010A), OutlookGatewayFailure.ComBusy)]
        [TestCase(unchecked((int)0x80070005), OutlookGatewayFailure.AccessDenied)]
        [TestCase(unchecked((int)0x80010108), OutlookGatewayFailure.NotReady)]
        [TestCase(unchecked((int)0x800401FD), OutlookGatewayFailure.NotReady)]
        [TestCase(unchecked((int)0x80004005), OutlookGatewayFailure.Internal)]
        public void MapsComFailuresWithoutLeakingProviderDetails(
            int hresult,
            OutlookGatewayFailure expected)
        {
            const string ProviderDetail = "mailbox and provider detail";
            var mapped = OutlookErrorMapper.Map(new TestComException(ProviderDetail, hresult));

            Assert.That(mapped.Failure, Is.EqualTo(expected));
            Assert.That(mapped.InnerException, Is.Null);
            Assert.That(mapped.Message, Does.Not.Contain(ProviderDetail));
            Assert.That(
                mapped.Message,
                Does.Not.Contain(hresult.ToString("X8", CultureInfo.InvariantCulture)));
        }

        [Test]
        public void PreservesExistingTypedGatewayFailure()
        {
            var expected = new OutlookGatewayException(OutlookGatewayFailure.Timeout);

            Assert.That(OutlookErrorMapper.Map(expected), Is.SameAs(expected));
        }

        [Test]
        public void UnexpectedFailureMapsToSafeInternalFailure()
        {
            var mapped = OutlookErrorMapper.Map(
                new InvalidCastException("sensitive failure detail"));

            Assert.That(mapped.Failure, Is.EqualTo(OutlookGatewayFailure.Internal));
            Assert.That(mapped.InnerException, Is.Null);
            Assert.That(mapped.Message, Is.EqualTo("The Outlook operation failed."));
        }

        private sealed class TestComException : COMException
        {
            public TestComException(string message, int hresult)
                : base(message, hresult)
            {
            }
        }
    }
}
#endif
