using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookClassicMcp.Transport
{
    public sealed class LoopbackHttpServer : IDisposable
    {
        private const int MaximumActiveHandlers = 4;
        private const string GenericRejectionBody =
            "{\"jsonrpc\":\"2.0\",\"error\":{\"code\":-32000,\"message\":\"Request rejected.\"},\"id\":null}";
        private const string ParseErrorBody =
            "{\"jsonrpc\":\"2.0\",\"error\":{\"code\":-32700,\"message\":\"Invalid JSON.\"},\"id\":null}";
        private const string InvalidRequestBody =
            "{\"jsonrpc\":\"2.0\",\"error\":{\"code\":-32600,\"message\":\"Invalid request.\"},\"id\":null}";
        private const string InternalErrorBody =
            "{\"jsonrpc\":\"2.0\",\"error\":{\"code\":-32603,\"message\":\"Internal error.\"},\"id\":null}";

        private static readonly UTF8Encoding Utf8 = new UTF8Encoding(false);
        private readonly object _gate = new object();
        private readonly HttpListener _listener;
        private readonly BearerToken _bearerToken;
        private readonly HttpRequestValidator _validator;
        private readonly McpRequestAdapter _adapter;
        private readonly TimeSpan _handlerDeadline;
        private readonly SemaphoreSlim _handlerSlots =
            new SemaphoreSlim(MaximumActiveHandlers, MaximumActiveHandlers);
        private readonly CancellationTokenSource _shutdown = new CancellationTokenSource();
        private readonly HashSet<Task> _handlerTasks = new HashSet<Task>();
        private readonly TaskCompletionSource<bool> _completion =
            new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        private Task _acceptLoop = Task.CompletedTask;
        private Task? _completionWorker;
        private bool _started;
        private bool _stopping;
        private bool _disposed;
        private int _activeHandlerCount;
        private int _shutdownRequested;

        public LoopbackHttpServer(
            BearerToken bearerToken,
            Func<OutlookStatusSnapshot> statusProvider)
            : this(bearerToken, statusProvider, RequestLimits.DefaultHandlerDeadline)
        {
        }

        internal LoopbackHttpServer(
            BearerToken bearerToken,
            Func<OutlookStatusSnapshot> statusProvider,
            TimeSpan handlerDeadline)
        {
            if (handlerDeadline <= TimeSpan.Zero ||
                handlerDeadline > RequestLimits.DefaultHandlerDeadline)
            {
                throw new ArgumentOutOfRangeException(nameof(handlerDeadline));
            }

            _bearerToken = bearerToken ?? throw new ArgumentNullException(nameof(bearerToken));
            _validator = new HttpRequestValidator(bearerToken);
            _adapter = new McpRequestAdapter(
                statusProvider ?? throw new ArgumentNullException(nameof(statusProvider)));
            _listener = new HttpListener();
            _listener.Prefixes.Add(LoopbackEndpoint.Prefix);
            _handlerDeadline = handlerDeadline;
        }

        public bool IsListening
        {
            get
            {
                lock (_gate)
                {
                    return _started && !_stopping && !_disposed && _listener.IsListening;
                }
            }
        }

        public int ActiveHandlerCount => Volatile.Read(ref _activeHandlerCount);

        public Task Completion => _completion.Task;

        public void Start()
        {
            lock (_gate)
            {
                ThrowIfUnavailable();
                if (_started)
                {
                    throw new InvalidOperationException("The loopback listener has already started.");
                }

                _listener.Start();
                _started = true;
                _acceptLoop = AcceptLoopAsync();
                _ = _acceptLoop.ContinueWith(
                    ObserveAcceptLoop,
                    CancellationToken.None,
                    TaskContinuationOptions.ExecuteSynchronously,
                    TaskScheduler.Default);
            }
        }

        public void Stop()
        {
            _ = BeginShutdown();
        }

        public Task BeginShutdown()
        {
            lock (_gate)
            {
                if (_stopping)
                {
                    return _completion.Task;
                }

                _stopping = true;
                Volatile.Write(ref _shutdownRequested, 1);
                Exception? startupFailure = null;
                try
                {
                    _listener.Close();
                }
                catch (ObjectDisposedException)
                {
                }
                catch (Exception exception) when (IsExpectedListenerShutdownException(exception))
                {
                }
                catch (Exception exception)
                {
                    startupFailure = exception;
                }

                _completionWorker = CompleteShutdownAsync(_acceptLoop, startupFailure);
                return _completion.Task;
            }
        }

        public void Dispose()
        {
            BeginShutdown().GetAwaiter().GetResult();
        }

        private async Task AcceptLoopAsync()
        {
            while (Volatile.Read(ref _shutdownRequested) == 0)
            {
                HttpListenerContext context;
                try
                {
                    context = await _listener.GetContextAsync().ConfigureAwait(false);
                }
                catch (Exception exception) when (IsExpectedListenerShutdownException(exception))
                {
                    return;
                }

                if (!_handlerSlots.Wait(0))
                {
                    try
                    {
                        await WriteErrorAndCloseAsync(
                            context.Response,
                            (int)HttpStatusCode.ServiceUnavailable,
                            GenericRejectionBody,
                            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                            {
                                ["Retry-After"] = "1",
                            },
                            CancellationToken.None).ConfigureAwait(false);
                    }
                    catch (Exception)
                    {
                        AbortContext(context);
                    }

                    continue;
                }

                Interlocked.Increment(ref _activeHandlerCount);
                var handlerTask = HandleContextAsync(context);
                lock (_gate)
                {
                    _handlerTasks.Add(handlerTask);
                }

                _ = handlerTask.ContinueWith(
                    RemoveCompletedHandler,
                    CancellationToken.None,
                    TaskContinuationOptions.ExecuteSynchronously,
                    TaskScheduler.Default);
            }
        }

        private async Task HandleContextAsync(HttpListenerContext context)
        {
            var response = context.Response;
            var mcpResponseStarted = false;
            using (var handlerCancellation =
                CancellationTokenSource.CreateLinkedTokenSource(_shutdown.Token))
            using (handlerCancellation.Token.Register(() => AbortContext(context)))
            {
                handlerCancellation.CancelAfter(_handlerDeadline);
                var handlerToken = handlerCancellation.Token;
                try
                {
                    var decision = _validator.Validate(CreateFacts(context.Request));
                    if (!decision.IsAccepted)
                    {
                        await WriteErrorAsync(
                            response,
                            decision.StatusCode,
                            GenericRejectionBody,
                            decision.RequiredHeaders,
                            handlerToken).ConfigureAwait(false);
                        return;
                    }

                    var readResult = await McpRequestAdapter.ReadSingleMessageAsync(
                        context.Request.InputStream,
                        (int)RequestLimits.MaximumRequestBodyBytes,
                        handlerToken).ConfigureAwait(false);
                    if (!readResult.Succeeded)
                    {
                        await WriteMessageReadFailureAsync(
                            response,
                            readResult.Failure,
                            handlerToken).ConfigureAwait(false);
                        return;
                    }

                    response.StatusCode = (int)HttpStatusCode.OK;
                    response.ContentType = "text/event-stream";
                    response.Headers["Cache-Control"] = "no-cache, no-store";
                    response.Headers["Content-Encoding"] = "identity";
                    response.SendChunked = true;
                    mcpResponseStarted = true;

                    var wroteResponse = await _adapter.HandleAsync(
                        readResult.Message!,
                        response.OutputStream,
                        handlerToken).ConfigureAwait(false);
                    if (!wroteResponse)
                    {
                        response.StatusCode = (int)HttpStatusCode.Accepted;
                        response.ContentType = null;
                        response.Headers.Remove("Content-Encoding");
                        response.Headers["Cache-Control"] = "no-store";
                        response.SendChunked = false;
                        response.ContentLength64 = 0;
                    }
                }
                catch (OperationCanceledException) when (handlerCancellation.IsCancellationRequested)
                {
                }
                catch (Exception exception) when (
                    IsExpectedRequestTerminationException(exception, handlerCancellation))
                {
                }
                catch (Exception)
                {
                    if (!mcpResponseStarted)
                    {
                        try
                        {
                            await WriteErrorAsync(
                                response,
                                (int)HttpStatusCode.InternalServerError,
                                InternalErrorBody,
                                null,
                                CancellationToken.None).ConfigureAwait(false);
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
                finally
                {
                    try
                    {
                        response.Close();
                    }
                    catch (ObjectDisposedException)
                    {
                    }
                    catch (Exception exception) when (
                        IsExpectedRequestTerminationException(exception, handlerCancellation))
                    {
                    }
                    finally
                    {
                        Interlocked.Decrement(ref _activeHandlerCount);
                        _handlerSlots.Release();
                    }

                }
            }
        }

        private static async Task WriteMessageReadFailureAsync(
            HttpListenerResponse response,
            McpMessageReadFailure failure,
            CancellationToken cancellationToken)
        {
            switch (failure)
            {
                case McpMessageReadFailure.PayloadTooLarge:
                    await WriteErrorAsync(
                        response,
                        (int)HttpStatusCode.RequestEntityTooLarge,
                        GenericRejectionBody,
                        null,
                        cancellationToken).ConfigureAwait(false);
                    return;
                case McpMessageReadFailure.BatchNotSupported:
                    await WriteErrorAsync(
                        response,
                        (int)HttpStatusCode.BadRequest,
                        InvalidRequestBody,
                        null,
                        cancellationToken).ConfigureAwait(false);
                    return;
                case McpMessageReadFailure.EmptyBody:
                case McpMessageReadFailure.MalformedJson:
                    await WriteErrorAsync(
                        response,
                        (int)HttpStatusCode.BadRequest,
                        ParseErrorBody,
                        null,
                        cancellationToken).ConfigureAwait(false);
                    return;
                default:
                    throw new ArgumentOutOfRangeException(nameof(failure));
            }
        }

        private static HttpRequestFacts CreateFacts(HttpListenerRequest request)
        {
            var headers = request.Headers.AllKeys
                .Where(name => name != null)
                .Select(name => new HttpHeaderFact(
                    name!,
                    request.Headers.GetValues(name!) ?? Array.Empty<string>()));
            return new HttpRequestFacts(
                request.RemoteEndPoint?.Address,
                request.RawUrl,
                request.HttpMethod,
                request.ContentLength64,
                headers);
        }

        private static Task WriteErrorAndCloseAsync(
            HttpListenerResponse response,
            int statusCode,
            string body,
            IReadOnlyDictionary<string, string>? requiredHeaders,
            CancellationToken cancellationToken)
        {
            return WriteAndCloseAsync(
                response,
                statusCode,
                body,
                requiredHeaders,
                cancellationToken);
        }

        private static async Task WriteAndCloseAsync(
            HttpListenerResponse response,
            int statusCode,
            string body,
            IReadOnlyDictionary<string, string>? requiredHeaders,
            CancellationToken cancellationToken)
        {
            try
            {
                await WriteErrorAsync(
                    response,
                    statusCode,
                    body,
                    requiredHeaders,
                    cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                response.Close();
            }
        }

        private static async Task WriteErrorAsync(
            HttpListenerResponse response,
            int statusCode,
            string body,
            IReadOnlyDictionary<string, string>? requiredHeaders,
            CancellationToken cancellationToken)
        {
            var bytes = Utf8.GetBytes(body);
            response.StatusCode = statusCode;
            response.ContentType = "application/json; charset=utf-8";
            response.Headers["Cache-Control"] = "no-store";
            if (requiredHeaders != null)
            {
                foreach (var header in requiredHeaders)
                {
                    response.AddHeader(header.Key, header.Value);
                }
            }

            response.ContentLength64 = bytes.Length;
            await response.OutputStream.WriteAsync(
                bytes,
                0,
                bytes.Length,
                cancellationToken).ConfigureAwait(false);
        }

        private void ObserveAcceptLoop(Task acceptLoop)
        {
            if (!acceptLoop.IsFaulted)
            {
                return;
            }

            _ = acceptLoop.Exception;
            _ = BeginShutdown();
        }

        private void RemoveCompletedHandler(Task handlerTask)
        {
            _ = handlerTask.Exception;
            lock (_gate)
            {
                _handlerTasks.Remove(handlerTask);
            }
        }

        private static void AbortContext(HttpListenerContext context)
        {
            try
            {
                context.Request.InputStream.Close();
            }
            catch (Exception)
            {
            }

            try
            {
                context.Response.Abort();
            }
            catch (Exception)
            {
            }
        }

        private async Task CompleteShutdownAsync(Task acceptLoop, Exception? startupFailure)
        {
            Exception? failure = startupFailure;
            try
            {
                await Task.Run(() => _shutdown.Cancel()).ConfigureAwait(false);
            }
            catch (Exception exception)
            {
                failure ??= exception;
            }

            try
            {
                await acceptLoop.ConfigureAwait(false);
            }
            catch (Exception exception)
            {
                failure ??= exception;
            }

            Task[] handlers;
            lock (_gate)
            {
                handlers = _handlerTasks.ToArray();
            }

            try
            {
                await Task.WhenAll(handlers).ConfigureAwait(false);
            }
            catch (Exception exception)
            {
                failure ??= exception;
            }

            try
            {
                _bearerToken.Dispose();
                _handlerSlots.Dispose();
                _shutdown.Dispose();
                lock (_gate)
                {
                    _disposed = true;
                }
            }
            catch (Exception exception)
            {
                failure ??= exception;
            }

            if (failure == null)
            {
                _completion.TrySetResult(true);
            }
            else
            {
                _completion.TrySetException(failure);
            }
        }

        private bool IsExpectedListenerShutdownException(Exception exception)
        {
            if (Volatile.Read(ref _shutdownRequested) == 0)
            {
                return false;
            }

            return exception is HttpListenerException ||
                exception is ObjectDisposedException ||
                exception is InvalidOperationException ||
                exception is OperationCanceledException ||
                (exception is ApplicationException &&
                    exception.HResult == unchecked((int)0x80070006));
        }

        private bool IsExpectedRequestTerminationException(
            Exception exception,
            CancellationTokenSource handlerCancellation)
        {
            if (!handlerCancellation.IsCancellationRequested &&
                Volatile.Read(ref _shutdownRequested) == 0)
            {
                return false;
            }

            return exception is HttpListenerException ||
                exception is ObjectDisposedException ||
                exception is InvalidOperationException ||
                exception is IOException ||
                exception is OperationCanceledException ||
                (exception is ApplicationException &&
                    exception.HResult == unchecked((int)0x80070006));
        }

        private void ThrowIfUnavailable()
        {
            if (_disposed || _stopping)
            {
                throw new ObjectDisposedException(nameof(LoopbackHttpServer));
            }
        }
    }
}
