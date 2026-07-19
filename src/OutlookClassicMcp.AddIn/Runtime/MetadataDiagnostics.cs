using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;

namespace OutlookClassicMcp.AddIn.Runtime
{
    internal enum RuntimeDiagnosticEvent
    {
        StartupCompleted,
        DependencyBindingCompleted,
        DispatcherProbeCompleted,
        HostQuiescent,
        ShutdownCompleted,
    }

    internal sealed class MetadataDiagnostics : IDisposable
    {
        private const long MaximumFileBytes = 5L * 1024L * 1024L;
        private const int MaximumFileCount = 5;
        private static readonly TimeSpan Retention = TimeSpan.FromDays(30);
        private static readonly Encoding LogEncoding = new UTF8Encoding(false);
        private readonly object _gate = new object();
        private readonly Guid _sessionId = Guid.NewGuid();
        private bool _enabled;
        private string? _logPath;

        private MetadataDiagnostics()
        {
        }

        public static MetadataDiagnostics Create()
        {
            var diagnostics = new MetadataDiagnostics();
            diagnostics.TryInitialize();
            return diagnostics;
        }

        public void RecordStartup(
            HostLifecycleState state,
            long durationTicks,
            bool debuggerAttached,
            int queueDepth,
            int trackedTaskCount)
        {
            Write(
                RuntimeDiagnosticEvent.StartupCompleted,
                state,
                true,
                durationTicks,
                queueDepth,
                trackedTaskCount,
                debuggerAttached: debuggerAttached,
                buildConfiguration: GetBuildConfiguration(),
                listenerActive: false);
        }

        public void RecordDependencyBinding(
            HostLifecycleState state,
            long durationTicks,
            Version coreVersion,
            int dependencyCount,
            string identitySha256,
            int queueDepth,
            int trackedTaskCount)
        {
            Write(
                RuntimeDiagnosticEvent.DependencyBindingCompleted,
                state,
                true,
                durationTicks,
                queueDepth,
                trackedTaskCount,
                coreVersion: coreVersion,
                dependencyCount: dependencyCount,
                dependencyIdentitySha256: identitySha256,
                listenerActive: false);
        }

        public void RecordDispatcherProbe(
            HostLifecycleState state,
            long durationTicks,
            OutlookThreadContext capturedThread,
            OutlookThreadContext executedThread,
            int queueDepth,
            int trackedTaskCount)
        {
            Write(
                RuntimeDiagnosticEvent.DispatcherProbeCompleted,
                state,
                true,
                durationTicks,
                queueDepth,
                trackedTaskCount,
                capturedThread: capturedThread,
                executedThread: executedThread,
                listenerActive: false);
        }

        public void RecordHostQuiescent(HostLifecycleState state, int queueDepth, int trackedTaskCount)
        {
            Write(
                RuntimeDiagnosticEvent.HostQuiescent,
                state,
                true,
                0,
                queueDepth,
                trackedTaskCount,
                listenerActive: false);
        }

        public void RecordShutdown(
            HostLifecycleState state,
            long durationTicks,
            int queueDepth,
            int trackedTaskCount)
        {
            Write(
                RuntimeDiagnosticEvent.ShutdownCompleted,
                state,
                true,
                durationTicks,
                queueDepth,
                trackedTaskCount,
                listenerActive: false);
        }

        public void RecordFailure(
            RuntimeDiagnosticEvent diagnosticEvent,
            HostLifecycleState state,
            long durationTicks,
            int queueDepth,
            int trackedTaskCount,
            Exception exception)
        {
            if (exception == null)
            {
                throw new ArgumentNullException(nameof(exception));
            }

            Write(
                diagnosticEvent,
                state,
                false,
                durationTicks,
                queueDepth,
                trackedTaskCount,
                listenerActive: false,
                exception: exception);
        }

        public void Dispose()
        {
            lock (_gate)
            {
                _enabled = false;
            }
        }

        private static string GetBuildConfiguration()
        {
#if DEBUG
            return "Debug";
#else
            return "Release";
#endif
        }

        private static string GetEventName(RuntimeDiagnosticEvent diagnosticEvent)
        {
            switch (diagnosticEvent)
            {
                case RuntimeDiagnosticEvent.StartupCompleted:
                    return "startup_completed";
                case RuntimeDiagnosticEvent.DependencyBindingCompleted:
                    return "dependency_binding_completed";
                case RuntimeDiagnosticEvent.DispatcherProbeCompleted:
                    return "dispatcher_probe_completed";
                case RuntimeDiagnosticEvent.HostQuiescent:
                    return "host_quiescent";
                case RuntimeDiagnosticEvent.ShutdownCompleted:
                    return "shutdown_completed";
                default:
                    throw new ArgumentOutOfRangeException(nameof(diagnosticEvent));
            }
        }

        private void TryInitialize()
        {
            try
            {
                var localApplicationData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                if (string.IsNullOrWhiteSpace(localApplicationData))
                {
                    return;
                }

                var directoryPath = Path.Combine(localApplicationData, "OutlookClassicMcp", "logs");
                var directory = Directory.CreateDirectory(directoryPath);
                ProtectDirectory(directory);
                RemoveExpiredLogs(directory);

                _logPath = Path.Combine(directory.FullName, "outlook-classic-mcp.log");
                ProtectLogFiles(directory, _logPath);
                _enabled = true;
            }
            catch
            {
                _enabled = false;
                _logPath = null;
            }
        }

        private static void ProtectDirectory(DirectoryInfo directory)
        {
            GetAllowedPrincipals(out var user, out var system);
            const InheritanceFlags inheritance =
                InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit;

            var security = new DirectorySecurity();
            security.SetAccessRuleProtection(true, false);
            security.SetOwner(user);
            security.AddAccessRule(new FileSystemAccessRule(
                user,
                FileSystemRights.FullControl,
                inheritance,
                PropagationFlags.None,
                AccessControlType.Allow));
            security.AddAccessRule(new FileSystemAccessRule(
                system,
                FileSystemRights.FullControl,
                inheritance,
                PropagationFlags.None,
                AccessControlType.Allow));
            directory.SetAccessControl(security);
        }

        private static void ProtectLogFiles(DirectoryInfo directory, string currentLogPath)
        {
            foreach (var file in directory.GetFiles("outlook-classic-mcp*.log", SearchOption.TopDirectoryOnly))
            {
                ProtectLogFile(file.FullName);
            }

            ProtectLogFile(currentLogPath);
        }

        private static void ProtectLogFile(string path)
        {
            using (var stream = new FileStream(
                path,
                FileMode.OpenOrCreate,
                FileAccess.Write,
                FileShare.Read))
            {
            }

            GetAllowedPrincipals(out var user, out var system);
            var security = new FileSecurity();
            security.SetAccessRuleProtection(true, false);
            security.SetOwner(user);
            security.AddAccessRule(new FileSystemAccessRule(
                user,
                FileSystemRights.FullControl,
                AccessControlType.Allow));
            security.AddAccessRule(new FileSystemAccessRule(
                system,
                FileSystemRights.FullControl,
                AccessControlType.Allow));
            File.SetAccessControl(path, security);
        }

        private static void GetAllowedPrincipals(
            out SecurityIdentifier user,
            out SecurityIdentifier system)
        {
            using (var identity = WindowsIdentity.GetCurrent())
            {
                user = identity.User
                    ?? throw new InvalidOperationException("The current Windows user has no security identifier.");
            }

            system = new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null);
        }

        private static void RemoveExpiredLogs(DirectoryInfo directory)
        {
            var cutoff = DateTime.UtcNow - Retention;
            foreach (var file in directory.GetFiles("outlook-classic-mcp*.log", SearchOption.TopDirectoryOnly))
            {
                if (file.LastWriteTimeUtc < cutoff)
                {
                    file.Delete();
                }
            }
        }

        private void Write(
            RuntimeDiagnosticEvent diagnosticEvent,
            HostLifecycleState state,
            bool succeeded,
            long durationTicks,
            int queueDepth,
            int trackedTaskCount,
            bool? debuggerAttached = null,
            string? buildConfiguration = null,
            Version? coreVersion = null,
            int? dependencyCount = null,
            string? dependencyIdentitySha256 = null,
            OutlookThreadContext? capturedThread = null,
            OutlookThreadContext? executedThread = null,
            bool? listenerActive = null,
            Exception? exception = null)
        {
            lock (_gate)
            {
                var logPath = _logPath;
                if (!_enabled || logPath == null || string.IsNullOrWhiteSpace(logPath))
                {
                    return;
                }

                try
                {
                    var line = CreateLine(
                        diagnosticEvent,
                        state,
                        succeeded,
                        durationTicks,
                        queueDepth,
                        trackedTaskCount,
                        debuggerAttached,
                        buildConfiguration,
                        coreVersion,
                        dependencyCount,
                        dependencyIdentitySha256,
                        capturedThread,
                        executedThread,
                        listenerActive,
                        exception);
                    RotateIfNeeded(logPath, LogEncoding.GetByteCount(line + Environment.NewLine));
                    ProtectLogFile(logPath);

                    using (var stream = new FileStream(
                        logPath,
                        FileMode.Append,
                        FileAccess.Write,
                        FileShare.Read,
                        4096,
                        FileOptions.SequentialScan))
                    using (var writer = new StreamWriter(stream, LogEncoding))
                    {
                        writer.WriteLine(line);
                    }
                }
                catch
                {
                    _enabled = false;
                }
            }
        }

        private string CreateLine(
            RuntimeDiagnosticEvent diagnosticEvent,
            HostLifecycleState state,
            bool succeeded,
            long durationTicks,
            int queueDepth,
            int trackedTaskCount,
            bool? debuggerAttached,
            string? buildConfiguration,
            Version? coreVersion,
            int? dependencyCount,
            string? dependencyIdentitySha256,
            OutlookThreadContext? capturedThread,
            OutlookThreadContext? executedThread,
            bool? listenerActive,
            Exception? exception)
        {
            var builder = new StringBuilder(512);
            Append(builder, "schema", "1");
            Append(builder, "timestamp", DateTimeOffset.UtcNow.ToString("O", CultureInfo.InvariantCulture));
            Append(builder, "session", _sessionId.ToString("N"));
            Append(builder, "event", GetEventName(diagnosticEvent));
            Append(builder, "state", state.ToString().ToLowerInvariant());
            Append(builder, "result", succeeded ? "success" : "failure");
            Append(builder, "duration_ticks", durationTicks.ToString(CultureInfo.InvariantCulture));
            Append(builder, "stopwatch_frequency", Stopwatch.Frequency.ToString(CultureInfo.InvariantCulture));
            Append(builder, "build", buildConfiguration ?? "-");
            Append(builder, "debugger_attached", FormatBoolean(debuggerAttached));
            Append(builder, "queue_depth", queueDepth.ToString(CultureInfo.InvariantCulture));
            Append(builder, "tracked_tasks", trackedTaskCount.ToString(CultureInfo.InvariantCulture));
            Append(builder, "listener_active", FormatBoolean(listenerActive));
            Append(builder, "mcp_core_version", coreVersion?.ToString() ?? "-");
            Append(builder, "dependency_count", FormatInteger(dependencyCount));
            Append(builder, "dependency_identity_sha256", dependencyIdentitySha256 ?? "-");
            Append(builder, "captured_managed_thread", FormatInteger(capturedThread?.ManagedThreadId));
            Append(builder, "captured_native_thread", FormatUnsignedInteger(capturedThread?.NativeThreadId));
            Append(builder, "captured_apartment", capturedThread?.ApartmentState.ToString() ?? "-");
            Append(builder, "executed_managed_thread", FormatInteger(executedThread?.ManagedThreadId));
            Append(builder, "executed_native_thread", FormatUnsignedInteger(executedThread?.NativeThreadId));
            Append(builder, "executed_apartment", executedThread?.ApartmentState.ToString() ?? "-");
            Append(builder, "exception_type", exception?.GetType().FullName ?? "-");
            Append(
                builder,
                "hresult",
                exception == null
                    ? "-"
                    : $"0x{exception.HResult.ToString("X8", CultureInfo.InvariantCulture)}");
            return builder.ToString();
        }

        private static void Append(StringBuilder builder, string key, string value)
        {
            if (builder.Length > 0)
            {
                builder.Append('\t');
            }

            builder.Append(key);
            builder.Append('=');
            builder.Append(value);
        }

        private static string FormatBoolean(bool? value)
        {
            return value.HasValue ? (value.Value ? "true" : "false") : "-";
        }

        private static string FormatInteger(int? value)
        {
            return value?.ToString(CultureInfo.InvariantCulture) ?? "-";
        }

        private static string FormatUnsignedInteger(uint? value)
        {
            return value?.ToString(CultureInfo.InvariantCulture) ?? "-";
        }

        private static void RotateIfNeeded(string currentPath, int pendingBytes)
        {
            var current = new FileInfo(currentPath);
            if (!current.Exists || current.Length + pendingBytes <= MaximumFileBytes)
            {
                return;
            }

            for (var index = MaximumFileCount - 1; index >= 1; index--)
            {
                var destination = GetRotatedPath(currentPath, index);
                if (File.Exists(destination))
                {
                    File.Delete(destination);
                }

                var source = index == 1 ? currentPath : GetRotatedPath(currentPath, index - 1);
                if (File.Exists(source))
                {
                    File.Move(source, destination);
                }
            }
        }

        private static string GetRotatedPath(string currentPath, int index)
        {
            var directory = Path.GetDirectoryName(currentPath)
                ?? throw new InvalidOperationException("The diagnostics path has no directory.");
            return Path.Combine(directory, $"outlook-classic-mcp.{index}.log");
        }
    }
}
