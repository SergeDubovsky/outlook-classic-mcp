using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using NUnit.Framework;
using OutlookClassicMcp.Core.Outlook;

namespace OutlookClassicMcp.Transport.Tests
{
    [TestFixture]
    [NonParallelizable]
    public sealed class LoopbackHttpIntegrationTests
    {
        private const string TokenText = "AAECAwQFBgcICQoLDA0ODxAREhMUFRYXGBkaGxwdHh8";
        private static readonly TimeSpan RequestTimeout = TimeSpan.FromSeconds(10);
        private static readonly HttpMethod[] UnsupportedMethods = { HttpMethod.Get, HttpMethod.Delete };
        private static readonly string[] AllowPost = { "POST" };
        private static readonly string[] NoStore = { "no-store" };
        private static readonly string[] BearerChallenge = { "Bearer" };
        private static readonly string[] RetryAfterOne = { "1" };
        private static readonly string[] SseCacheDirectives = { "no-cache", "no-store" };
        private static readonly string[] IdentityEncoding = { "identity" };
        private static readonly string[] ValidTokenAuthorizationHeaders =
            { "Bearer " + TokenText };
        private static readonly string CursorQueryHash = new string('a', 64);
        private static readonly string[] StatusDataProperties =
            { "hostState", "listenerReady", "version", "readDiagnostics" };
        private static readonly string[] PhaseFourToolNames =
        {
            "outlook_status",
            "outlook_probe",
            "outlook_list_mailboxes",
            "outlook_list_folders",
            "outlook_list_messages",
            "outlook_search_messages",
            "outlook_get_message",
            "outlook_get_conversation",
            "outlook_list_attachments",
        };

        [Test]
        public async Task UnauthorizedResponsesAreIndistinguishableAcrossRawForms()
        {
            using (var server = CreateServer(TokenText))
            using (var client = CreateHttpClient())
            {
                server.Start();
                var missing = await SendJsonAsync(client, PingRequest(1), authorization: null);
                var wrong = await SendJsonAsync(client, PingRequest(2), "Bearer " + CreateToken(101));
                var malformed = new[]
                {
                    "Basic " + TokenText,
                    "Bearer  " + TokenText,
                    "Bearer",
                    "Bearer " + TokenText + "=",
                };

                AssertUnauthorized(missing);
                AssertSameResponse(missing, wrong);
                foreach (var authorization in malformed)
                {
                    var response = await SendJsonAsync(client, PingRequest(3), authorization);
                    AssertSameResponse(missing, response);
                }
            }
        }

        [Test]
        public async Task OriginAndUnsupportedMethodsFailBeforeMcpProcessing()
        {
            using (var server = CreateServer(TokenText))
            using (var client = CreateHttpClient())
            {
                server.Start();
                var origin = await SendJsonAsync(
                    client,
                    PingRequest(10),
                    ValidAuthorization(TokenText),
                    origin: "https://example.test");
                Assert.That(origin.StatusCode, Is.EqualTo(HttpStatusCode.Forbidden));
                Assert.That(origin.Body, Does.Contain("Request rejected."));
                Assert.That(origin.CorsHeaders, Is.Empty);

                foreach (var method in UnsupportedMethods)
                {
                    var response = await SendJsonAsync(
                        client,
                        PingRequest(11),
                        ValidAuthorization(TokenText),
                        method: method);
                    Assert.That(response.StatusCode, Is.EqualTo(HttpStatusCode.MethodNotAllowed));
                    Assert.That(response.Allow, Is.EqualTo(AllowPost));
                }
            }
        }

        [Test]
        public async Task MalformedBatchAndOversizedBodiesReturnBoundedErrors()
        {
            var providerCalls = 0;
            using (var server = CreateServer(TokenText, () =>
            {
                providerCalls++;
                return new OutlookStatusSnapshot("online", listenerReady: true, "1.0.0");
            }))
            using (var client = CreateHttpClient())
            {
                server.Start();
                var malformed = await SendJsonAsync(
                    client,
                    "{not-json",
                    ValidAuthorization(TokenText));
                Assert.That(malformed.StatusCode, Is.EqualTo(HttpStatusCode.BadRequest));
                Assert.That(malformed.Body, Does.Contain("Invalid JSON."));

                var batch = await SendJsonAsync(
                    client,
                    "[" + StatusRequest(12) + "]",
                    ValidAuthorization(TokenText));
                Assert.That(batch.StatusCode, Is.EqualTo(HttpStatusCode.BadRequest));
                Assert.That(batch.Body, Does.Contain("Invalid request."));
                Assert.That(providerCalls, Is.Zero);

                var oversizedBody = new string(' ', (int)RequestLimits.MaximumRequestBodyBytes + 1);
                var oversized = await SendJsonAsync(
                    client,
                    oversizedBody,
                    ValidAuthorization(TokenText));
                Assert.That(oversized.StatusCode, Is.EqualTo(HttpStatusCode.RequestEntityTooLarge));
                Assert.That(Encoding.UTF8.GetByteCount(oversized.Body), Is.LessThan(256));

                var recovery = await SendJsonAsync(
                    client,
                    PingRequest(13),
                    ValidAuthorization(TokenText));
                AssertSseResponse(recovery, expectedId: 13);
            }
        }

        [Test]
        public async Task MediaNegotiationEncodingAndProtocolErrorsAreBounded()
        {
            using (var server = CreateServer(TokenText))
            using (var client = CreateHttpClient())
            {
                server.Start();
                var missingContentType = await SendJsonAsync(
                    client,
                    PingRequest(14),
                    ValidAuthorization(TokenText),
                    contentType: null);
                Assert.That(missingContentType.StatusCode, Is.EqualTo(HttpStatusCode.UnsupportedMediaType));

                var missingAccept = await SendJsonAsync(
                    client,
                    PingRequest(15),
                    ValidAuthorization(TokenText),
                    accept: null);
                Assert.That(missingAccept.StatusCode, Is.EqualTo(HttpStatusCode.NotAcceptable));

                var unsupportedVersion = await SendJsonAsync(
                    client,
                    PingRequest(16),
                    ValidAuthorization(TokenText),
                    protocolVersion: "2099-01-01");
                Assert.That(unsupportedVersion.StatusCode, Is.EqualTo(HttpStatusCode.BadRequest));

                var encoded = await SendJsonAsync(
                    client,
                    PingRequest(17),
                    ValidAuthorization(TokenText),
                    contentEncoding: "gzip");
                Assert.That(encoded.StatusCode, Is.EqualTo(HttpStatusCode.UnsupportedMediaType));

                var optionalParameters = await SendJsonAsync(
                    client,
                    PingRequest(18),
                    ValidAuthorization(TokenText),
                    contentType: "Application/Json; Charset=UTF-8; profile=example",
                    accept: "text/event-stream; q=0.5, application/json; q=1");
                AssertSseResponse(optionalParameters, expectedId: 18);
            }
        }

        [Test]
        public async Task MissingProtocolVersionUsesTheBackwardCompatibilityPath()
        {
            using (var server = CreateServer(TokenText))
            using (var client = CreateHttpClient())
            {
                server.Start();
                var response = await SendJsonAsync(
                    client,
                    PingRequest(19),
                    ValidAuthorization(TokenText),
                    protocolVersion: null);

                AssertSseResponse(response, expectedId: 19);
            }
        }

        [Test]
        public async Task RequestUsesExactSseContractAndNotificationUsesEmpty202Contract()
        {
            using (var server = CreateServer(TokenText))
            using (var client = CreateHttpClient())
            {
                server.Start();
                var initialize = await SendJsonAsync(
                    client,
                    InitializeRequest(20),
                    ValidAuthorization(TokenText),
                    protocolVersion: null);
                AssertSseResponse(initialize, expectedId: 20);
                using (var payload = ParseSingleSsePayload(initialize.Body))
                {
                    var result = payload.RootElement.GetProperty("result");
                    Assert.That(result.GetProperty("protocolVersion").GetString(), Is.EqualTo("2025-11-25"));
                    Assert.That(result.GetProperty("serverInfo").GetProperty("name").GetString(),
                        Is.EqualTo("outlook-classic-mcp"));
                }

                var notification = await SendJsonAsync(
                    client,
                    "{\"jsonrpc\":\"2.0\",\"method\":\"notifications/initialized\",\"params\":{}}",
                    ValidAuthorization(TokenText));
                Assert.That(notification.StatusCode, Is.EqualTo(HttpStatusCode.Accepted));
                Assert.That(notification.Body, Is.Empty);
                Assert.That(notification.ContentType, Is.Null);
                Assert.That(notification.ContentEncodings, Is.Empty);
                Assert.That(notification.CacheControl, Is.EqualTo(NoStore));
            }
        }

        [TestCase("2024-11-05")]
        [TestCase("2025-03-26")]
        [TestCase("2025-06-18")]
        [TestCase("2025-11-25")]
        public async Task EveryPinnedProtocolVersionNegotiatesThroughTheSdk(string protocolVersion)
        {
            using (var server = CreateServer(TokenText))
            using (var client = CreateHttpClient())
            {
                server.Start();
                var initialize = await SendJsonAsync(
                    client,
                    InitializeRequest(22, protocolVersion),
                    ValidAuthorization(TokenText),
                    protocolVersion: protocolVersion);
                AssertSseResponse(initialize, expectedId: 22);
                using (var payload = ParseSingleSsePayload(initialize.Body))
                {
                    Assert.That(
                        payload.RootElement.GetProperty("result").GetProperty("protocolVersion").GetString(),
                        Is.EqualTo(protocolVersion));
                }
            }
        }

        [Test]
        public async Task ToolListStatusAndReadExposeBoundedPhaseFourData()
        {
            var gateway = new FakeOutlookGateway();
            var mailbox = CreateLoopbackMailbox();
            var calls = 0;
            gateway.ListMailboxesHandler = (_, __) =>
            {
                calls++;
                return Task.FromResult(calls == 1
                    ? new OutlookMailboxPage(
                        new[] { mailbox },
                        new OutlookMailboxKeysetAnchor(
                            mailbox.DisplayName,
                            mailbox.Mailbox))
                    : new OutlookMailboxPage(
                        Array.Empty<OutlookMailboxSummary>(),
                        null));
            };

            using (var server = CreateServer(
                TokenText,
                () => new OutlookStatusSnapshot(
                    "online",
                    listenerReady: true,
                    "1.0.0",
                    new OutlookReadDiagnosticsSnapshot(21, 18, 3, 6, 50)),
                gateway))
            using (var client = CreateHttpClient())
            {
                server.Start();
                var list = await SendJsonAsync(
                    client,
                    "{\"jsonrpc\":\"2.0\",\"id\":30,\"method\":\"tools/list\",\"params\":{}}",
                    ValidAuthorization(TokenText));
                AssertSseResponse(list, expectedId: 30);
                using (var payload = ParseSingleSsePayload(list.Body))
                {
                    var tools = payload.RootElement.GetProperty("result").GetProperty("tools");
                    Assert.That(tools.GetArrayLength(), Is.EqualTo(PhaseFourToolNames.Length));
                    for (var index = 0; index < PhaseFourToolNames.Length; index++)
                    {
                        Assert.That(
                            tools[index].GetProperty("name").GetString(),
                            Is.EqualTo(PhaseFourToolNames[index]));
                    }
                }

                var status = await SendJsonAsync(
                    client,
                    "{\"jsonrpc\":\"2.0\",\"id\":31,\"method\":\"tools/call\",\"params\":{\"name\":\"outlook_status\",\"arguments\":{}}}",
                    ValidAuthorization(TokenText));
                AssertSseResponse(status, expectedId: 31);
                using (var payload = ParseSingleSsePayload(status.Body))
                {
                    var result = payload.RootElement.GetProperty("result");
                    Assert.That(result.GetProperty("isError").GetBoolean(), Is.False);
                    var structured = result.GetProperty("structuredContent");
                    Assert.That(structured.GetProperty("ok").GetBoolean(), Is.True);
                    Assert.That(structured.GetProperty("partial").GetBoolean(), Is.False);
                    Assert.That(structured.GetProperty("warnings").GetArrayLength(), Is.Zero);
                    Assert.That(structured.GetProperty("operationId").GetString(),
                        Does.Match("^[0-9a-f]{32}$"));
                    var data = structured.GetProperty("data");
                    Assert.That(data.GetProperty("hostState").GetString(), Is.EqualTo("online"));
                    Assert.That(data.GetProperty("listenerReady").GetBoolean(), Is.True);
                    Assert.That(data.GetProperty("version").GetString(), Is.EqualTo("1.0.0"));
                    var diagnostics = data.GetProperty("readDiagnostics");
                    Assert.That(diagnostics.GetProperty("comAcquired").GetInt64(), Is.EqualTo(21));
                    Assert.That(diagnostics.GetProperty("comReleased").GetInt64(), Is.EqualTo(18));
                    Assert.That(diagnostics.GetProperty("comOutstanding").GetInt64(), Is.EqualTo(3));
                    Assert.That(diagnostics.GetProperty("comPeak").GetInt64(), Is.EqualTo(6));
                    Assert.That(
                        diagnostics.GetProperty("materializedItemHighWater").GetInt64(),
                        Is.EqualTo(50));
                    Assert.That(data.EnumerateObject().Select(property => property.Name),
                        Is.EquivalentTo(StatusDataProperties));
                }

                var firstRead = await SendJsonAsync(
                    client,
                    "{\"jsonrpc\":\"2.0\",\"id\":32,\"method\":\"tools/call\",\"params\":{" +
                    "\"name\":\"outlook_list_mailboxes\",\"arguments\":{}}}",
                    ValidAuthorization(TokenText));
                AssertSseResponse(firstRead, expectedId: 32);
                string cursor;
                using (var payload = ParseSingleSsePayload(firstRead.Body))
                {
                    var result = payload.RootElement.GetProperty("result");
                    Assert.That(result.GetProperty("isError").GetBoolean(), Is.False);
                    var structured = result.GetProperty("structuredContent");
                    Assert.That(structured.GetProperty("ok").GetBoolean(), Is.True);
                    var data = structured.GetProperty("data");
                    Assert.That(data.GetProperty("mailboxes").GetArrayLength(), Is.EqualTo(1));
                    cursor = data.GetProperty("nextCursor").GetString()!;
                }

                var secondRead = await SendJsonAsync(
                    client,
                    "{\"jsonrpc\":\"2.0\",\"id\":33,\"method\":\"tools/call\",\"params\":{" +
                    "\"name\":\"outlook_list_mailboxes\",\"arguments\":{" +
                    "\"cursor\":\"" + cursor + "\"}}}",
                    ValidAuthorization(TokenText));
                AssertSseResponse(secondRead, expectedId: 33);
                using (var payload = ParseSingleSsePayload(secondRead.Body))
                {
                    Assert.That(
                        payload.RootElement.GetProperty("result").GetProperty("isError").GetBoolean(),
                        Is.False);
                }

                Assert.That(gateway.ListMailboxesCallCount, Is.EqualTo(2));
                Assert.That(gateway.LastListMailboxesRequest!.Anchor, Is.Not.Null);
                Assert.That(
                    gateway.LastListMailboxesRequest.Anchor!.Mailbox.StoreId,
                    Is.EqualTo(mailbox.Mailbox.StoreId));
            }
        }

        [Test]
        public async Task TwentySequentialProbeCallsPreserveIdsAndKeepTheListenerHealthy()
        {
            var gateway = new FakeOutlookGateway();
            using (var server = CreateServer(
                TokenText,
                () => new OutlookStatusSnapshot("online", listenerReady: true, "1.0.0"),
                gateway))
            using (var client = CreateHttpClient())
            {
                server.Start();
                for (var index = 0; index < 20; index++)
                {
                    var id = 40 + index;
                    var response = await SendJsonAsync(
                        client,
                        ProbeRequest(id),
                        ValidAuthorization(TokenText));
                    AssertSseResponse(response, id);
                    using (var payload = ParseSingleSsePayload(response.Body))
                    {
                        var result = payload.RootElement.GetProperty("result");
                        Assert.That(result.GetProperty("isError").GetBoolean(), Is.False);
                        Assert.That(
                            result.GetProperty("structuredContent").GetProperty("data")
                                .GetProperty("dispatcher").GetProperty("matchesCapturedThread")
                                .GetBoolean(),
                            Is.True);
                    }
                }

                Assert.That(gateway.CallCount, Is.EqualTo(20));
                Assert.That(server.IsListening, Is.True);
                Assert.That(server.ActiveHandlerCount, Is.Zero);
            }
        }

        [Test]
        public async Task FourConcurrentProbeCallsAreBoundedAndComplete()
        {
            using (var entered = new CountdownEvent(4))
            {
                var release = new TaskCompletionSource<bool>(
                    TaskCreationOptions.RunContinuationsAsynchronously);
                var activeCalls = 0;
                var maximumActiveCalls = 0;
                var gateway = new FakeOutlookGateway(async cancellationToken =>
                {
                    var active = Interlocked.Increment(ref activeCalls);
                    UpdateMaximum(ref maximumActiveCalls, active);
                    entered.Signal();
                    try
                    {
                        await release.Task;
                        cancellationToken.ThrowIfCancellationRequested();
                        return ProbeTestData.CreateSnapshot();
                    }
                    finally
                    {
                        Interlocked.Decrement(ref activeCalls);
                    }
                });
                using (var server = CreateServer(
                    TokenText,
                    () => new OutlookStatusSnapshot("online", listenerReady: true, "1.0.0"),
                    gateway))
                using (var client = CreateHttpClient())
                {
                    server.Start();
                    var calls = Enumerable.Range(0, 4)
                        .Select(index => SendJsonAsync(
                            client,
                            ProbeRequest(70 + index),
                            ValidAuthorization(TokenText)))
                        .ToArray();
                    try
                    {
                        Assert.That(entered.Wait(RequestTimeout), Is.True);
                    }
                    finally
                    {
                        release.TrySetResult(true);
                    }

                    var responses = await Task.WhenAll(calls);
                    for (var index = 0; index < responses.Length; index++)
                    {
                        AssertSseResponse(responses[index], 70 + index);
                    }

                    Assert.That(maximumActiveCalls, Is.EqualTo(4));
                    Assert.That(server.ActiveHandlerCount, Is.Zero);
                }
            }
        }

        [Test]
        public async Task ProbeFailureIsAProtocolSuccessAndTheNextProbeRecovers()
        {
            var calls = 0;
            var gateway = new FakeOutlookGateway(_ =>
            {
                calls++;
                return calls == 1
                    ? Task.FromException<OutlookProbeSnapshot>(
                        new OutlookGatewayException(OutlookGatewayFailure.ComBusy))
                    : Task.FromResult(ProbeTestData.CreateSnapshot());
            });
            using (var server = CreateServer(
                TokenText,
                () => new OutlookStatusSnapshot("online", listenerReady: true, "1.0.0"),
                gateway))
            using (var client = CreateHttpClient())
            {
                server.Start();
                var failed = await SendJsonAsync(
                    client,
                    ProbeRequest(80),
                    ValidAuthorization(TokenText));
                AssertSseResponse(failed, 80);
                using (var payload = ParseSingleSsePayload(failed.Body))
                {
                    var result = payload.RootElement.GetProperty("result");
                    Assert.That(result.GetProperty("isError").GetBoolean(), Is.True);
                    Assert.That(
                        result.GetProperty("structuredContent").GetProperty("error")
                            .GetProperty("code").GetString(),
                        Is.EqualTo("COM_BUSY"));
                }

                var recovered = await SendJsonAsync(
                    client,
                    ProbeRequest(81),
                    ValidAuthorization(TokenText));
                AssertSseResponse(recovered, 81);
                using (var payload = ParseSingleSsePayload(recovered.Body))
                {
                    Assert.That(
                        payload.RootElement.GetProperty("result").GetProperty("isError").GetBoolean(),
                        Is.False);
                }

                Assert.That(server.IsListening, Is.True);
                Assert.That(gateway.CallCount, Is.EqualTo(2));
            }
        }

        [Test]
        public async Task InnerProbeDeadlineReturnsTimeoutAndObservesAnIgnoringGatewayBeforeRecovery()
        {
            var pending = new TaskCompletionSource<OutlookProbeSnapshot>(
                TaskCreationOptions.RunContinuationsAsynchronously);
            var calls = 0;
            var gateway = new FakeOutlookGateway(_ =>
            {
                var call = Interlocked.Increment(ref calls);
                return call == 1
                    ? pending.Task
                    : Task.FromResult(ProbeTestData.CreateSnapshot());
            });
            Assert.That(BearerToken.TryCreate(TokenText, out var token), Is.True);
            using (var server = new LoopbackHttpServer(
                token,
                () => new OutlookStatusSnapshot("online", listenerReady: true, "1.0.0"),
                gateway,
                TimeSpan.FromMilliseconds(150),
                TimeSpan.FromSeconds(2)))
            using (var client = CreateHttpClient())
            {
                server.Start();
                var timedOut = await SendJsonAsync(
                    client,
                    ProbeRequest(82),
                    ValidAuthorization(TokenText));
                AssertSseResponse(timedOut, 82);
                using (var payload = ParseSingleSsePayload(timedOut.Body))
                {
                    var result = payload.RootElement.GetProperty("result");
                    Assert.That(result.GetProperty("isError").GetBoolean(), Is.True);
                    var error = result.GetProperty("structuredContent").GetProperty("error");
                    Assert.That(error.GetProperty("code").GetString(), Is.EqualTo("TIMEOUT"));
                    Assert.That(error.GetProperty("retryable").GetBoolean(), Is.True);
                }

                Assert.That(pending.Task.IsCompleted, Is.False);
                Assert.That(gateway.LastCancellationToken.IsCancellationRequested, Is.True);
                pending.TrySetException(new InvalidOperationException(
                    "eventual sensitive gateway detail"));

                var recovered = await SendJsonAsync(
                    client,
                    ProbeRequest(83),
                    ValidAuthorization(TokenText));
                AssertSseResponse(recovered, 83);
                using (var payload = ParseSingleSsePayload(recovered.Body))
                {
                    Assert.That(
                        payload.RootElement.GetProperty("result").GetProperty("isError").GetBoolean(),
                        Is.False);
                    Assert.That(payload.RootElement.GetRawText(),
                        Does.Not.Contain("eventual sensitive gateway detail"));
                }

                Assert.That(server.IsListening, Is.True);
                Assert.That(gateway.CallCount, Is.EqualTo(2));
            }
        }

        [Test]
        public async Task TwentyUnauthorizedThenTwentyValidRequestsDoNotWedgeTheListener()
        {
            using (var server = CreateServer(TokenText))
            using (var client = CreateHttpClient())
            {
                server.Start();
                for (var index = 0; index < 20; index++)
                {
                    var unauthorized = index % 2 == 0;
                    var rejected = await SendJsonAsync(
                        client,
                        unauthorized ? PingRequest(100 + index) : "{not-json",
                        unauthorized
                            ? "Bearer " + CreateToken(index + 120)
                            : ValidAuthorization(TokenText));
                    Assert.That(
                        rejected.StatusCode,
                        Is.EqualTo(unauthorized ? HttpStatusCode.Unauthorized : HttpStatusCode.BadRequest));
                }

                for (var index = 0; index < 20; index++)
                {
                    var id = 200 + index;
                    var accepted = await SendJsonAsync(
                        client,
                        PingRequest(id),
                        ValidAuthorization(TokenText));
                    AssertSseResponse(accepted, id);
                }

                Assert.That(server.IsListening, Is.True);
                Assert.That(server.ActiveHandlerCount, Is.Zero);
            }
        }

        [Test]
        public async Task ShutdownStopsAdmissionBeforeWaitingForAnActiveHandler()
        {
            using (var entered = new ManualResetEventSlim(false))
            using (var release = new ManualResetEventSlim(false))
            using (var server = CreateServer(TokenText, () =>
            {
                entered.Set();
                release.Wait(RequestTimeout);
                return new OutlookStatusSnapshot("online", listenerReady: true, "1.0.0");
            }))
            using (var client = CreateHttpClient())
            {
                var ownedToken = GetPrivateField<BearerToken>(server, "_bearerToken");
                var ownedCodec = GetPrivateField<HmacCursorCodec>(server, "_cursorCodec");
                var cursorPayload = new MailboxCursorPayload(
                    CursorQueryHash,
                    "Mailbox",
                    "store-a");
                server.Start();
                var activeRequest = SendJsonAsync(
                    client,
                    StatusRequest(240),
                    ValidAuthorization(TokenText));
                Assert.That(entered.Wait(RequestTimeout), Is.True);

                var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                var completion = server.BeginShutdown();
                Assert.That(server.IsListening, Is.False);
                Assert.That(completion.IsCompleted, Is.False);
                Assert.That(
                    ownedToken.MatchesAuthorizationHeaders(ValidTokenAuthorizationHeaders),
                    Is.True);
                Assert.DoesNotThrow((Action)(() => _ = ownedCodec.Encode(cursorPayload)));
                using (var replacement = CreateServer(CreateToken(33)))
                {
                    Assert.DoesNotThrow((Action)replacement.Start);
                    Assert.That(stopwatch.Elapsed, Is.LessThan(TimeSpan.FromSeconds(3)));
                }

                release.Set();
                try
                {
                    _ = await activeRequest;
                }
                catch (HttpRequestException)
                {
                }
                catch (TaskCanceledException)
                {
                }

                await completion;
                Assert.That(server.ActiveHandlerCount, Is.Zero);
                Assert.Throws<ObjectDisposedException>((Action)(() =>
                    ownedToken.MatchesAuthorizationHeaders(ValidTokenAuthorizationHeaders)));
                Assert.Throws<ObjectDisposedException>((Action)(() =>
                    _ = ownedCodec.Encode(cursorPayload)));
            }
        }

        [Test]
        public async Task FifthConcurrentHandlerIsRejectedAndCapacityRecovers()
        {
            using (var entered = new CountdownEvent(4))
            using (var release = new ManualResetEventSlim(false))
            using (var server = CreateServer(TokenText, () =>
            {
                entered.Signal();
                release.Wait(RequestTimeout);
                return new OutlookStatusSnapshot("online", listenerReady: true, "1.0.0");
            }))
            using (var client = CreateHttpClient())
            {
                server.Start();
                var held = Enumerable.Range(0, 4)
                    .Select(index => SendJsonAsync(
                        client,
                        StatusRequest(250 + index),
                        ValidAuthorization(TokenText)))
                    .ToArray();

                var allEntered = entered.Wait(RequestTimeout);
                HttpResult? rejected = null;
                try
                {
                    if (allEntered)
                    {
                        rejected = await SendJsonAsync(
                            client,
                            PingRequest(254),
                            ValidAuthorization(TokenText));
                    }
                }
                finally
                {
                    release.Set();
                }

                var completed = await Task.WhenAll(held);
                Assert.That(allEntered, Is.True);
                Assert.That(rejected, Is.Not.Null);
                Assert.That(rejected!.StatusCode, Is.EqualTo(HttpStatusCode.ServiceUnavailable));
                Assert.That(rejected.RetryAfter, Is.EqualTo(RetryAfterOne));
                for (var index = 0; index < completed.Length; index++)
                {
                    AssertSseResponse(completed[index], 250 + index);
                }

                var recovered = await SendJsonAsync(
                    client,
                    PingRequest(255),
                    ValidAuthorization(TokenText));
                AssertSseResponse(recovered, expectedId: 255);
            }
        }

        [Test]
        public async Task StalledAuthenticatedBodiesExpireAndCapacityRecovers()
        {
            Assert.That(BearerToken.TryCreate(TokenText, out var token), Is.True);
            using (var server = new LoopbackHttpServer(
                token,
                () => new OutlookStatusSnapshot("online", listenerReady: true, "1.0.0"),
                new FakeOutlookGateway(),
                TimeSpan.FromSeconds(1),
                TimeSpan.FromSeconds(2)))
            using (var client = CreateHttpClient())
            {
                var stalledClients = new List<TcpClient>();
                try
                {
                    server.Start();
                    var requestBytes = Encoding.ASCII.GetBytes(
                        "POST /mcp/ HTTP/1.1\r\n" +
                        "Host: 127.0.0.1:8765\r\n" +
                        "Authorization: Bearer " + TokenText + "\r\n" +
                        "Content-Type: application/json\r\n" +
                        "Accept: application/json, text/event-stream\r\n" +
                        "MCP-Protocol-Version: 2025-11-25\r\n" +
                        "Transfer-Encoding: chunked\r\n\r\n");
                    for (var index = 0; index < 4; index++)
                    {
                        var stalledClient = new TcpClient();
                        stalledClients.Add(stalledClient);
                        await stalledClient.ConnectAsync(IPAddress.Loopback, 8765);
                        await stalledClient.GetStream().WriteAsync(
                            requestBytes,
                            0,
                            requestBytes.Length);
                        await stalledClient.GetStream().FlushAsync();
                    }

                    Assert.That(
                        await WaitForConditionAsync(
                            () => server.ActiveHandlerCount == 4,
                            TimeSpan.FromSeconds(3)),
                        Is.True);
                    Assert.That(
                        await WaitForConditionAsync(
                            () => server.ActiveHandlerCount == 0,
                            TimeSpan.FromSeconds(5)),
                        Is.True);

                    var recovered = await SendJsonAsync(
                        client,
                        PingRequest(256),
                        ValidAuthorization(TokenText));
                    AssertSseResponse(recovered, expectedId: 256);
                }
                finally
                {
                    foreach (var stalledClient in stalledClients)
                    {
                        stalledClient.Dispose();
                    }
                }
            }
        }

        [Test]
        public async Task TokenRotationAndRestartRetireTheOldTokenAndReleaseThePort()
        {
            var replacementToken = CreateToken(64);
            using (var client = CreateHttpClient())
            {
                using (var first = CreateServer(TokenText))
                {
                    first.Start();
                    var accepted = await SendJsonAsync(
                        client,
                        PingRequest(300),
                        ValidAuthorization(TokenText));
                    AssertSseResponse(accepted, expectedId: 300);
                }

                using (var second = CreateServer(replacementToken))
                {
                    Assert.DoesNotThrow((Action)second.Start);
                    var retired = await SendJsonAsync(
                        client,
                        PingRequest(301),
                        ValidAuthorization(TokenText));
                    Assert.That(retired.StatusCode, Is.EqualTo(HttpStatusCode.Unauthorized));

                    var replacement = await SendJsonAsync(
                        client,
                        PingRequest(302),
                        ValidAuthorization(replacementToken));
                    AssertSseResponse(replacement, expectedId: 302);
                }

                using (var third = CreateServer(TokenText))
                {
                    Assert.DoesNotThrow((Action)third.Start);
                    Assert.That(third.IsListening, Is.True);
                }
            }
        }

        [Test]
        public async Task ProcessTokenRotationTakesEffectOnlyForANewServer()
        {
            var replacementToken = CreateToken(72);
            var original = Environment.GetEnvironmentVariable(
                BearerToken.EnvironmentVariableName,
                EnvironmentVariableTarget.Process);
            try
            {
                Environment.SetEnvironmentVariable(
                    BearerToken.EnvironmentVariableName,
                    TokenText,
                    EnvironmentVariableTarget.Process);
                using (var firstToken = BearerToken.LoadFromProcessEnvironment())
                using (var first = new LoopbackHttpServer(
                    firstToken,
                    () => new OutlookStatusSnapshot("online", true, "1.0.0"),
                    new FakeOutlookGateway()))
                using (var client = CreateHttpClient())
                {
                    first.Start();
                    Environment.SetEnvironmentVariable(
                        BearerToken.EnvironmentVariableName,
                        replacementToken,
                        EnvironmentVariableTarget.Process);

                    var retained = await SendJsonAsync(
                        client,
                        PingRequest(303),
                        ValidAuthorization(TokenText));
                    AssertSseResponse(retained, expectedId: 303);
                    var premature = await SendJsonAsync(
                        client,
                        PingRequest(304),
                        ValidAuthorization(replacementToken));
                    Assert.That(premature.StatusCode, Is.EqualTo(HttpStatusCode.Unauthorized));
                }

                using (var secondToken = BearerToken.LoadFromProcessEnvironment())
                using (var second = new LoopbackHttpServer(
                    secondToken,
                    () => new OutlookStatusSnapshot("online", true, "1.0.0"),
                    new FakeOutlookGateway()))
                using (var client = CreateHttpClient())
                {
                    second.Start();
                    var retired = await SendJsonAsync(
                        client,
                        PingRequest(305),
                        ValidAuthorization(TokenText));
                    Assert.That(retired.StatusCode, Is.EqualTo(HttpStatusCode.Unauthorized));
                    var replaced = await SendJsonAsync(
                        client,
                        PingRequest(306),
                        ValidAuthorization(replacementToken));
                    AssertSseResponse(replaced, expectedId: 306);
                }
            }
            finally
            {
                Environment.SetEnvironmentVariable(
                    BearerToken.EnvironmentVariableName,
                    original,
                    EnvironmentVariableTarget.Process);
            }
        }

        [Test]
        public async Task PinnedMcpClientInitializesPingsListsAndCallsPhaseFourTools()
        {
            using (var server = CreateServer(TokenText))
            using (var httpClient = CreateHttpClient())
            {
                server.Start();
                var transport = new HttpClientTransport(
                    new HttpClientTransportOptions
                    {
                        Endpoint = LoopbackEndpoint.Address,
                        TransportMode = HttpTransportMode.StreamableHttp,
                        Name = "loopback-integration-test",
                        ConnectionTimeout = RequestTimeout,
                        OwnsSession = false,
                        AdditionalHeaders = new Dictionary<string, string>
                        {
                            ["Authorization"] = ValidAuthorization(TokenText),
                        },
                    },
                    httpClient,
                    loggerFactory: null,
                    ownsHttpClient: false);
                McpClient? client = null;
                try
                {
                    using (var cancellation = new CancellationTokenSource(RequestTimeout))
                    {
                        client = await McpClient.CreateAsync(
                            transport,
                            clientOptions: null,
                            loggerFactory: null,
                            cancellation.Token);
                        var ping = await client.PingAsync(new PingRequestParams(), cancellation.Token);
                        Assert.That(ping, Is.Not.Null);

                        var tools = await client.ListToolsAsync(
                            new ListToolsRequestParams(),
                            cancellation.Token);
                        Assert.That(tools.Tools.Select(tool => tool.Name),
                            Is.EqualTo(PhaseFourToolNames));

                        var status = await client.CallToolAsync(
                            new CallToolRequestParams
                            {
                                Name = "outlook_status",
                                Arguments = new Dictionary<string, JsonElement>(),
                            },
                            cancellation.Token);
                        Assert.That(status.IsError, Is.False);
                        Assert.That(status.StructuredContent.HasValue, Is.True);
                        Assert.That(
                            status.StructuredContent!.Value.GetProperty("data")
                                .GetProperty("hostState").GetString(),
                            Is.EqualTo("online"));

                        var probe = await client.CallToolAsync(
                            new CallToolRequestParams
                            {
                                Name = "outlook_probe",
                                Arguments = new Dictionary<string, JsonElement>(),
                            },
                            cancellation.Token);
                        Assert.That(probe.IsError, Is.False);
                        Assert.That(probe.StructuredContent.HasValue, Is.True);
                        Assert.That(
                            probe.StructuredContent!.Value.GetProperty("data")
                                .GetProperty("dispatcher").GetProperty("matchesCapturedThread")
                                .GetBoolean(),
                            Is.True);
                    }
                }
                finally
                {
                    if (client != null)
                    {
                        await client.DisposeAsync();
                    }
                    else
                    {
                        await transport.DisposeAsync();
                    }
                }
            }
        }

        private static LoopbackHttpServer CreateServer(string encodedToken)
        {
            return CreateServer(
                encodedToken,
                () => new OutlookStatusSnapshot("online", listenerReady: true, "1.0.0"));
        }

        private static LoopbackHttpServer CreateServer(
            string encodedToken,
            Func<OutlookStatusSnapshot> statusProvider)
        {
            return CreateServer(encodedToken, statusProvider, new FakeOutlookGateway());
        }

        private static LoopbackHttpServer CreateServer(
            string encodedToken,
            Func<OutlookStatusSnapshot> statusProvider,
            IOutlookGateway outlookGateway)
        {
            Assert.That(BearerToken.TryCreate(encodedToken, out var token), Is.True);
            return new LoopbackHttpServer(token, statusProvider, outlookGateway);
        }

        private static OutlookMailboxSummary CreateLoopbackMailbox()
        {
            var mailbox = new MailboxRef("loopback-store");
            return new OutlookMailboxSummary(
                mailbox,
                "Loopback mailbox",
                OutlookStoreType.PrimaryExchangeMailbox,
                new OutlookStoreCapabilities(true, false, true),
                new OutlookStandardFolderReferences(
                    new FolderRef(mailbox.StoreId, "loopback-inbox"),
                    null,
                    null,
                    null,
                    null));
        }

        private static T GetPrivateField<T>(object instance, string fieldName)
            where T : class
        {
            var field = instance.GetType().GetField(
                fieldName,
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.NonPublic);
            Assert.That(field, Is.Not.Null);
            return (T)field!.GetValue(instance)!;
        }

        private static HttpClient CreateHttpClient()
        {
            var handler = new HttpClientHandler
            {
                UseProxy = false,
                AllowAutoRedirect = false,
                AutomaticDecompression = DecompressionMethods.None,
                MaxConnectionsPerServer = 16,
            };
            return new HttpClient(handler, disposeHandler: true)
            {
                Timeout = RequestTimeout,
            };
        }

        private static async Task<HttpResult> SendJsonAsync(
            HttpClient client,
            string body,
            string? authorization,
            HttpMethod? method = null,
            string? origin = null,
            string? protocolVersion = "2025-11-25",
            string? contentType = "application/json; charset=utf-8",
            string? accept = "application/json, text/event-stream",
            string? contentEncoding = null)
        {
            using (var request = new HttpRequestMessage(method ?? HttpMethod.Post, LoopbackEndpoint.Address))
            using (var cancellation = new CancellationTokenSource(RequestTimeout))
            {
                if (request.Method == HttpMethod.Post)
                {
                    request.Content = new StringContent(body, Encoding.UTF8);
                    request.Content.Headers.Remove("Content-Type");
                    if (contentType != null)
                    {
                        request.Content.Headers.TryAddWithoutValidation("Content-Type", contentType);
                    }
                    if (contentEncoding != null)
                    {
                        request.Content.Headers.TryAddWithoutValidation("Content-Encoding", contentEncoding);
                    }
                }

                request.Headers.Host = HttpRequestValidator.CanonicalHost;
                if (accept != null)
                {
                    request.Headers.TryAddWithoutValidation("Accept", accept);
                }
                if (authorization != null)
                {
                    request.Headers.TryAddWithoutValidation("Authorization", authorization);
                }
                if (origin != null)
                {
                    request.Headers.TryAddWithoutValidation("Origin", origin);
                }
                if (protocolVersion != null)
                {
                    request.Headers.TryAddWithoutValidation("MCP-Protocol-Version", protocolVersion);
                }

                using (var response = await client.SendAsync(
                    request,
                    HttpCompletionOption.ResponseContentRead,
                    cancellation.Token))
                {
                    return new HttpResult(
                        response.StatusCode,
                        await response.Content.ReadAsStringAsync(),
                        response.Content.Headers.ContentType?.ToString(),
                        response.Content.Headers.ContentEncoding.ToArray(),
                        GetHeaderValues(response, "Cache-Control"),
                        GetHeaderValues(response, "WWW-Authenticate"),
                        GetHeaderValues(response, "Allow"),
                        GetHeaderValues(response, "Retry-After"),
                        response.Headers.TransferEncodingChunked == true,
                        response.Headers
                            .Where(header => header.Key.StartsWith("Access-Control-", StringComparison.OrdinalIgnoreCase))
                            .Select(header => header.Key)
                            .ToArray());
                }
            }
        }

        private static string[] GetHeaderValues(HttpResponseMessage response, string name)
        {
            if (response.Headers.TryGetValues(name, out var responseValues))
            {
                return responseValues.ToArray();
            }
            if (response.Content.Headers.TryGetValues(name, out var contentValues))
            {
                return contentValues.ToArray();
            }

            return Array.Empty<string>();
        }

        private static void AssertUnauthorized(HttpResult response)
        {
            Assert.Multiple((Action)(() =>
            {
                Assert.That(response.StatusCode, Is.EqualTo(HttpStatusCode.Unauthorized));
                Assert.That(response.ContentType, Is.EqualTo("application/json; charset=utf-8"));
                Assert.That(response.CacheControl, Is.EqualTo(NoStore));
                Assert.That(response.WwwAuthenticate, Is.EqualTo(BearerChallenge));
                Assert.That(response.Body, Does.Not.Contain("outlook_status"));
            }));
        }

        private static void AssertSameResponse(HttpResult expected, HttpResult actual)
        {
            Assert.Multiple((Action)(() =>
            {
                Assert.That(actual.StatusCode, Is.EqualTo(expected.StatusCode));
                Assert.That(actual.Body, Is.EqualTo(expected.Body));
                Assert.That(actual.ContentType, Is.EqualTo(expected.ContentType));
                Assert.That(actual.ContentEncodings, Is.EqualTo(expected.ContentEncodings));
                Assert.That(actual.CacheControl, Is.EqualTo(expected.CacheControl));
                Assert.That(actual.WwwAuthenticate, Is.EqualTo(expected.WwwAuthenticate));
            }));
        }

        private static void AssertSseResponse(HttpResult response, int expectedId)
        {
            Assert.Multiple((Action)(() =>
            {
                Assert.That(response.StatusCode, Is.EqualTo(HttpStatusCode.OK));
                Assert.That(response.ContentType, Is.EqualTo("text/event-stream"));
                Assert.That(
                    response.CacheControl
                        .SelectMany(value => value.Split(','))
                        .Select(value => value.Trim()),
                    Is.EquivalentTo(SseCacheDirectives));
                Assert.That(response.ContentEncodings, Is.EqualTo(IdentityEncoding));
                Assert.That(response.TransferEncodingChunked, Is.True);
            }));

            using (var payload = ParseSingleSsePayload(response.Body))
            {
                Assert.That(payload.RootElement.GetProperty("id").GetInt32(), Is.EqualTo(expectedId));
            }
        }

        private static JsonDocument ParseSingleSsePayload(string body)
        {
            var normalized = body.Replace("\r\n", "\n");
            var lines = normalized.Split('\n');
            Assert.That(lines.Length, Is.GreaterThanOrEqualTo(4));
            Assert.That(lines[0], Is.EqualTo("event: message"));
            Assert.That(lines[1], Does.StartWith("data: "));
            Assert.That(lines.Skip(2), Is.All.Empty);
            return JsonDocument.Parse(lines[1].Substring("data: ".Length));
        }

        private static string PingRequest(int id)
        {
            return $"{{\"jsonrpc\":\"2.0\",\"id\":{id},\"method\":\"ping\",\"params\":{{}}}}";
        }

        private static string InitializeRequest(int id, string protocolVersion = "2025-11-25")
        {
            return $"{{\"jsonrpc\":\"2.0\",\"id\":{id},\"method\":\"initialize\",\"params\":{{" +
                $"\"protocolVersion\":\"{protocolVersion}\",\"capabilities\":{{}}," +
                "\"clientInfo\":{\"name\":\"loopback-integration-test\",\"version\":\"1.0\"}}}";
        }

        private static string StatusRequest(int id)
        {
            return $"{{\"jsonrpc\":\"2.0\",\"id\":{id},\"method\":\"tools/call\",\"params\":{{" +
                "\"name\":\"outlook_status\",\"arguments\":{}}}";
        }

        private static string ProbeRequest(int id)
        {
            return $"{{\"jsonrpc\":\"2.0\",\"id\":{id},\"method\":\"tools/call\",\"params\":{{" +
                "\"name\":\"outlook_probe\",\"arguments\":{}}}";
        }

        private static string ValidAuthorization(string token)
        {
            return "Bearer " + token;
        }

        private static string CreateToken(int seed)
        {
            var bytes = Enumerable.Range(0, 32)
                .Select(index => unchecked((byte)(seed + index)))
                .ToArray();
            return Convert.ToBase64String(bytes)
                .TrimEnd('=')
                .Replace('+', '-')
                .Replace('/', '_');
        }

        private static async Task<bool> WaitForConditionAsync(
            Func<bool> condition,
            TimeSpan timeout)
        {
            var deadline = DateTime.UtcNow + timeout;
            while (DateTime.UtcNow < deadline)
            {
                if (condition())
                {
                    return true;
                }

                await Task.Delay(25);
            }

            return condition();
        }

        private static void UpdateMaximum(ref int maximum, int value)
        {
            var current = Volatile.Read(ref maximum);
            while (value > current)
            {
                var observed = Interlocked.CompareExchange(ref maximum, value, current);
                if (observed == current)
                {
                    return;
                }

                current = observed;
            }
        }

        private sealed class HttpResult
        {
            public HttpResult(
                HttpStatusCode statusCode,
                string body,
                string? contentType,
                string[] contentEncodings,
                string[] cacheControl,
                string[] wwwAuthenticate,
                string[] allow,
                string[] retryAfter,
                bool transferEncodingChunked,
                string[] corsHeaders)
            {
                StatusCode = statusCode;
                Body = body;
                ContentType = contentType;
                ContentEncodings = contentEncodings;
                CacheControl = cacheControl;
                WwwAuthenticate = wwwAuthenticate;
                Allow = allow;
                RetryAfter = retryAfter;
                TransferEncodingChunked = transferEncodingChunked;
                CorsHeaders = corsHeaders;
            }

            public HttpStatusCode StatusCode { get; }

            public string Body { get; }

            public string? ContentType { get; }

            public string[] ContentEncodings { get; }

            public string[] CacheControl { get; }

            public string[] WwwAuthenticate { get; }

            public string[] Allow { get; }

            public string[] RetryAfter { get; }

            public bool TransferEncodingChunked { get; }

            public string[] CorsHeaders { get; }
        }
    }
}
