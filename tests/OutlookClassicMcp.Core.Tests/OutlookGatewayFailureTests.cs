using System;
using System.Linq;
using NUnit.Framework;
using OutlookClassicMcp.Core.Outlook;

namespace OutlookClassicMcp.Core.Tests
{
    [TestFixture]
    public sealed class OutlookGatewayFailureTests
    {
        [TestCase(OutlookGatewayFailure.NotReady, "Outlook is not ready.")]
        [TestCase(OutlookGatewayFailure.Degraded, "The Outlook integration is degraded.")]
        [TestCase(OutlookGatewayFailure.Stopping, "The Outlook integration is stopping.")]
        [TestCase(OutlookGatewayFailure.QueueFull, "The Outlook request queue is full.")]
        [TestCase(OutlookGatewayFailure.Timeout, "The Outlook request timed out.")]
        [TestCase(OutlookGatewayFailure.ComBusy, "Outlook is busy.")]
        [TestCase(OutlookGatewayFailure.AccessDenied, "Outlook denied access.")]
        [TestCase(OutlookGatewayFailure.ObjectModelGuard, "Outlook blocked the operation.")]
        [TestCase(OutlookGatewayFailure.StaDispatchFailed, "The Outlook UI thread dispatch failed.")]
        [TestCase(OutlookGatewayFailure.Internal, "The Outlook operation failed.")]
        public void FailureHasFixedSafeMessage(OutlookGatewayFailure failure, string expectedMessage)
        {
            var exception = new OutlookGatewayException(failure);

            AssertAll(() =>
            {
                Assert.That(exception.Failure, Is.EqualTo(failure));
                Assert.That(exception.Message, Is.EqualTo(expectedMessage));
                Assert.That(exception.Message, Does.Not.Contain("HRESULT"));
                Assert.That(exception.Message, Does.Not.Contain("0x"));
                Assert.That(exception.Message.Any(char.IsControl), Is.False);
                Assert.That(exception.InnerException, Is.Null);
            });
        }

        [Test]
        public void UndefinedFailureIsRejectedWithoutProducingAnExceptionMessage()
        {
            AssertThrows<ArgumentOutOfRangeException>(
                () => new OutlookGatewayException((OutlookGatewayFailure)99));
        }

        [Test]
        public void GatewayExceptionDoesNotAcceptArbitraryMessagesOrInnerExceptions()
        {
            var publicConstructors = typeof(OutlookGatewayException).GetConstructors();

            Assert.That(publicConstructors, Has.Length.EqualTo(1));
            Assert.That(
                publicConstructors[0].GetParameters().Select(parameter => parameter.ParameterType),
                Is.EqualTo(new[] { typeof(OutlookGatewayFailure) }));
        }

        [Test]
        public void GatewayFailureSetIsClosedAndStable()
        {
            Assert.That(
                typeof(OutlookGatewayFailure).GetEnumNames(),
                Is.EqualTo(new[]
                {
                    nameof(OutlookGatewayFailure.NotReady),
                    nameof(OutlookGatewayFailure.Degraded),
                    nameof(OutlookGatewayFailure.Stopping),
                    nameof(OutlookGatewayFailure.QueueFull),
                    nameof(OutlookGatewayFailure.Timeout),
                    nameof(OutlookGatewayFailure.ComBusy),
                    nameof(OutlookGatewayFailure.AccessDenied),
                    nameof(OutlookGatewayFailure.ObjectModelGuard),
                    nameof(OutlookGatewayFailure.StaDispatchFailed),
                    nameof(OutlookGatewayFailure.Internal),
                }));
        }

        private static void AssertAll(Action assertions)
        {
            Assert.Multiple((Action)(() => assertions()));
        }

        private static void AssertThrows<TException>(Func<object?> action)
            where TException : Exception
        {
            Assert.Catch<TException>((Action)(() => _ = action()));
        }

    }
}
