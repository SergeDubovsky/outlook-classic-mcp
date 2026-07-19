using System;

namespace OutlookClassicMcp.AddIn.Runtime
{
    internal sealed class HostLifecycle
    {
        private readonly object _gate = new object();
        private HostLifecycleState _state = HostLifecycleState.Created;

        public HostLifecycleState State
        {
            get
            {
                lock (_gate)
                {
                    return _state;
                }
            }
        }

        public bool TryBeginStartup()
        {
            lock (_gate)
            {
                if (_state != HostLifecycleState.Created &&
                    _state != HostLifecycleState.Degraded &&
                    _state != HostLifecycleState.Paused)
                {
                    return false;
                }

                _state = HostLifecycleState.Starting;
                return true;
            }
        }

        public bool TryMarkOnline()
        {
            return TryTransition(HostLifecycleState.Starting, HostLifecycleState.Online);
        }

        public bool TryMarkDegraded()
        {
            return TryTransition(HostLifecycleState.Starting, HostLifecycleState.Degraded);
        }

        public bool TryBeginPause()
        {
            lock (_gate)
            {
                if (_state != HostLifecycleState.Online)
                {
                    return false;
                }

                _state = HostLifecycleState.Pausing;
                return true;
            }
        }

        public bool TryMarkPaused()
        {
            return TryTransition(HostLifecycleState.Pausing, HostLifecycleState.Paused);
        }

        public bool TryBeginStop()
        {
            lock (_gate)
            {
                if (_state == HostLifecycleState.Stopping || _state == HostLifecycleState.Stopped)
                {
                    return false;
                }

                _state = HostLifecycleState.Stopping;
                return true;
            }
        }

        public void MarkStopped()
        {
            lock (_gate)
            {
                if (_state == HostLifecycleState.Stopped)
                {
                    return;
                }

                if (_state != HostLifecycleState.Stopping)
                {
                    throw CreateInvalidTransitionException(_state, HostLifecycleState.Stopped);
                }

                _state = HostLifecycleState.Stopped;
            }
        }

        private bool TryTransition(HostLifecycleState expected, HostLifecycleState next)
        {
            lock (_gate)
            {
                if (_state != expected)
                {
                    return false;
                }

                _state = next;
                return true;
            }
        }

        private static InvalidOperationException CreateInvalidTransitionException(
            HostLifecycleState current,
            HostLifecycleState next)
        {
            return new InvalidOperationException($"Cannot transition the host from {current} to {next}.");
        }
    }
}
