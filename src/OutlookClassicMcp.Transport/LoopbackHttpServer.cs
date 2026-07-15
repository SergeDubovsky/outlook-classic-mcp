using System;
using System.Net;

namespace OutlookClassicMcp.Transport
{
    /// <summary>
    /// Phase 0 listener-lifecycle proof. Request processing is added in Phase 2.
    /// </summary>
    public sealed class LoopbackHttpServer : IDisposable
    {
        private readonly HttpListener _listener;
        private bool _disposed;

        public LoopbackHttpServer()
        {
            _listener = new HttpListener();
            _listener.Prefixes.Add(LoopbackEndpoint.Prefix);
        }

        public bool IsListening => !_disposed && _listener.IsListening;

        public void Start()
        {
            ThrowIfDisposed();
            _listener.Start();
        }

        public void Stop()
        {
            if (!_disposed && _listener.IsListening)
            {
                _listener.Stop();
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            _listener.Close();
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(LoopbackHttpServer));
            }
        }
    }
}
