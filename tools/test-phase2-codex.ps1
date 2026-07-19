#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateRange(15, 180)]
    [int]$TimeoutSeconds = 90
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$tokenVariable = 'OUTLOOK_MCP_TOKEN'
$token = [Environment]::GetEnvironmentVariable($tokenVariable, 'Process')
if ($token -notmatch '^[A-Za-z0-9_-]{43}$') {
    $token = [Environment]::GetEnvironmentVariable($tokenVariable, 'User')
}
if ($token -notmatch '^[A-Za-z0-9_-]{43}$') {
    throw "$tokenVariable is not a canonical 32-byte base64url token in process or current-user scope."
}
if (-not (Get-Command codex -CommandType Application -ErrorAction SilentlyContinue)) {
    throw 'The Codex CLI is not available on PATH.'
}

if ($null -eq ('OutlookClassicMcp.Tools.ProcessJob' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookClassicMcp.Tools
{
    public sealed class ProcessJob : IDisposable
    {
        private const uint KillOnJobClose = 0x00002000;
        private IntPtr handle;

        public ProcessJob()
        {
            handle = CreateJobObject(IntPtr.Zero, null);
            if (handle == IntPtr.Zero)
            {
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }

            var limits = new JobObjectExtendedLimitInformation();
            limits.BasicLimitInformation.LimitFlags = KillOnJobClose;
            var size = Marshal.SizeOf(typeof(JobObjectExtendedLimitInformation));
            var pointer = Marshal.AllocHGlobal(size);
            try
            {
                Marshal.StructureToPtr(limits, pointer, false);
                if (!SetInformationJobObject(handle, 9, pointer, (uint)size))
                {
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                }
            }
            catch
            {
                CloseHandle(handle);
                handle = IntPtr.Zero;
                throw;
            }
            finally
            {
                Marshal.FreeHGlobal(pointer);
            }
        }

        public void Add(Process process)
        {
            if (process == null)
            {
                throw new ArgumentNullException("process");
            }
            if (handle == IntPtr.Zero)
            {
                throw new ObjectDisposedException("ProcessJob");
            }
            if (!AssignProcessToJobObject(handle, process.Handle))
            {
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }
        }

        public void TerminateAndVerify(int timeoutMilliseconds)
        {
            if (timeoutMilliseconds <= 0)
            {
                throw new ArgumentOutOfRangeException("timeoutMilliseconds");
            }
            if (handle == IntPtr.Zero)
            {
                throw new ObjectDisposedException("ProcessJob");
            }
            if (!TerminateJobObject(handle, 1))
            {
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }

            var stopwatch = Stopwatch.StartNew();
            while (GetActiveProcessCount() != 0 && stopwatch.ElapsedMilliseconds < timeoutMilliseconds)
            {
                Thread.Sleep(25);
            }
            if (GetActiveProcessCount() != 0)
            {
                throw new TimeoutException("The process job did not become empty before the deadline.");
            }
        }

        public void Dispose()
        {
            if (handle != IntPtr.Zero)
            {
                if (!CloseHandle(handle))
                {
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                }
                handle = IntPtr.Zero;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct IoCounters
        {
            public ulong ReadOperationCount;
            public ulong WriteOperationCount;
            public ulong OtherOperationCount;
            public ulong ReadTransferCount;
            public ulong WriteTransferCount;
            public ulong OtherTransferCount;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct JobObjectBasicLimitInformation
        {
            public long PerProcessUserTimeLimit;
            public long PerJobUserTimeLimit;
            public uint LimitFlags;
            public UIntPtr MinimumWorkingSetSize;
            public UIntPtr MaximumWorkingSetSize;
            public uint ActiveProcessLimit;
            public UIntPtr Affinity;
            public uint PriorityClass;
            public uint SchedulingClass;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct JobObjectExtendedLimitInformation
        {
            public JobObjectBasicLimitInformation BasicLimitInformation;
            public IoCounters IoInfo;
            public UIntPtr ProcessMemoryLimit;
            public UIntPtr JobMemoryLimit;
            public UIntPtr PeakProcessMemoryUsed;
            public UIntPtr PeakJobMemoryUsed;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct JobObjectBasicAccountingInformation
        {
            public long TotalUserTime;
            public long TotalKernelTime;
            public long ThisPeriodTotalUserTime;
            public long ThisPeriodTotalKernelTime;
            public uint TotalPageFaultCount;
            public uint TotalProcesses;
            public uint ActiveProcesses;
            public uint TotalTerminatedProcesses;
        }

        private uint GetActiveProcessCount()
        {
            var size = Marshal.SizeOf(typeof(JobObjectBasicAccountingInformation));
            var pointer = Marshal.AllocHGlobal(size);
            try
            {
                uint returnedLength;
                if (!QueryInformationJobObject(handle, 1, pointer, (uint)size, out returnedLength))
                {
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                }
                var accounting = (JobObjectBasicAccountingInformation)Marshal.PtrToStructure(
                    pointer,
                    typeof(JobObjectBasicAccountingInformation));
                return accounting.ActiveProcesses;
            }
            finally
            {
                Marshal.FreeHGlobal(pointer);
            }
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateJobObject(IntPtr jobAttributes, string name);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetInformationJobObject(
            IntPtr job,
            int informationClass,
            IntPtr information,
            uint informationLength);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool TerminateJobObject(IntPtr job, uint exitCode);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool QueryInformationJobObject(
            IntPtr job,
            int informationClass,
            IntPtr information,
            uint informationLength,
            out uint returnLength);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr handle);
    }

    public sealed class BoundedTextCapture
    {
        private readonly StringBuilder buffer;
        private readonly int maximumCharacters;

        public BoundedTextCapture(TextReader reader, int maximumCharacters)
        {
            if (reader == null)
            {
                throw new ArgumentNullException("reader");
            }
            if (maximumCharacters <= 0)
            {
                throw new ArgumentOutOfRangeException("maximumCharacters");
            }

            this.maximumCharacters = maximumCharacters;
            buffer = new StringBuilder(Math.Min(maximumCharacters, 8192));
            Completion = CaptureAsync(reader);
        }

        public Task Completion { get; private set; }

        public bool LimitExceeded { get; private set; }

        public string Text
        {
            get
            {
                if (!Completion.IsCompleted)
                {
                    throw new InvalidOperationException("Capture is not complete.");
                }
                return buffer.ToString();
            }
        }

        private async Task CaptureAsync(TextReader reader)
        {
            var characters = new char[4096];
            int count;
            while ((count = await reader.ReadAsync(characters, 0, characters.Length).ConfigureAwait(false)) != 0)
            {
                var remaining = maximumCharacters - buffer.Length;
                if (remaining > 0)
                {
                    buffer.Append(characters, 0, Math.Min(remaining, count));
                }
                if (count > remaining)
                {
                    LimitExceeded = true;
                }
            }
        }
    }
}
'@
}

function Get-OptionalProperty {
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($null -eq $InputObject) {
        return $null
    }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }
    return $property.Value
}

function Get-PropertyEntry {
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory = $true)][string[]]$Names
    )

    if ($null -eq $InputObject) {
        return $null
    }
    foreach ($name in $Names) {
        $property = $InputObject.PSObject.Properties[$name]
        if ($null -ne $property) {
            return $property
        }
    }
    return $null
}

function Assert-ExactProperties {
    param(
        [Parameter(Mandatory = $true)][object]$InputObject,
        [Parameter(Mandatory = $true)][string[]]$Expected,
        [Parameter(Mandatory = $true)][string]$Context
    )

    $actual = @($InputObject.PSObject.Properties.Name | Sort-Object)
    $wanted = @($Expected | Sort-Object)
    if (($actual -join "`n") -cne ($wanted -join "`n")) {
        throw "$Context did not contain the exact expected fields."
    }
}

function Protect-DiagnosticText {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) {
        return ''
    }
    $protected = ([string]$Value).Replace($token, '[redacted]')
    $protected = $protected -replace '(?<![A-Za-z0-9_-])[A-Za-z0-9_-]{43}(?![A-Za-z0-9_-])', '[redacted-token]'
    $protected = $protected -replace '[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '?'
    if ($protected.Length -gt 1024) {
        return $protected.Substring(0, 1024)
    }
    return $protected
}

function Wait-ForOutputTasks {
    param(
        [Parameter(Mandatory = $true)][OutlookClassicMcp.Tools.BoundedTextCapture]$StandardOutput,
        [Parameter(Mandatory = $true)][OutlookClassicMcp.Tools.BoundedTextCapture]$StandardError
    )

    $tasks = [System.Threading.Tasks.Task[]]@($StandardOutput.Completion, $StandardError.Completion)
    if (-not [System.Threading.Tasks.Task]::WaitAll($tasks, 5000)) {
        throw 'Codex output streams did not close within five seconds after its process tree stopped.'
    }
}

$previousToken = [Environment]::GetEnvironmentVariable($tokenVariable, 'Process')
$process = $null
$processJob = $null
$launchGate = $null
$stdoutCapture = $null
$stderrCapture = $null
$processTreeClosed = $false

try {
    [Environment]::SetEnvironmentVariable($tokenVariable, $token, 'Process')

    $codexCommand = @(Get-Command codex -CommandType Application)[0]
    $powerShellCommand = @(Get-Command powershell.exe -CommandType Application)[0]
    $launchGateName = "Local\OutlookClassicMcp-Codex-$([Guid]::NewGuid().ToString('N'))"
    $launchGate = [Threading.EventWaitHandle]::new(
        $false,
        [Threading.EventResetMode]::ManualReset,
        $launchGateName)
    $escapedGateName = $launchGateName.Replace("'", "''")
    $escapedCodexPath = $codexCommand.Source.Replace("'", "''")
    $helperScript = @"
`$gate = [Threading.EventWaitHandle]::OpenExisting('$escapedGateName')
try {
    if (-not `$gate.WaitOne(30000)) { exit 124 }
}
finally {
    `$gate.Dispose()
}
& '$escapedCodexPath' 'exec' '-c' 'skills.include_instructions=false' '--disable' 'multi_agent' '--ephemeral' '--sandbox' 'read-only' '--strict-config' '--ignore-rules' '--json' '-'
exit `$LASTEXITCODE
"@
    $encodedHelper = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($helperScript))
    $startInfo = [Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $powerShellCommand.Source
    $startInfo.Arguments = "-NoLogo -NoProfile -NonInteractive -EncodedCommand $encodedHelper"
    $startInfo.WorkingDirectory = $repositoryRoot
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardInput = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true

    $processJob = [OutlookClassicMcp.Tools.ProcessJob]::new()
    $process = [Diagnostics.Process]::Start($startInfo)
    if ($null -eq $process) {
        throw 'Codex did not return a process for the Phase 2 probe.'
    }
    $processJob.Add($process)

    $stdoutCapture = [OutlookClassicMcp.Tools.BoundedTextCapture]::new(
        $process.StandardOutput,
        1024 * 1024)
    $stderrCapture = [OutlookClassicMcp.Tools.BoundedTextCapture]::new(
        $process.StandardError,
        256 * 1024)
    $prompt = @'
Use the outlook_classic MCP server and call outlook_status exactly once. Do not run shell commands, inspect files, browse the web, or call any other tool. Verify that the native structured content has ok=true, data.hostState="online", and data.listenerReady=true. If and only if all checks pass, reply with exactly MCP_PHASE2_CODEX_OK.
'@
    $process.StandardInput.WriteLine($prompt.Trim())
    $process.StandardInput.Close()
    $null = $launchGate.Set()

    $completedBeforeDeadline = $process.WaitForExit($TimeoutSeconds * 1000)
    $processJob.TerminateAndVerify(5000)
    $processJob.Dispose()
    $processJob = $null
    if (-not $process.WaitForExit(5000)) {
        throw 'The Codex process remained alive after its process-tree job was closed.'
    }
    Wait-ForOutputTasks -StandardOutput $stdoutCapture -StandardError $stderrCapture
    $processTreeClosed = $true

    if ($stdoutCapture.LimitExceeded -or $stderrCapture.LimitExceeded) {
        throw 'Codex output exceeded the bounded Phase 2 probe capture limit.'
    }
    $stdout = $stdoutCapture.Text
    $stderr = $stderrCapture.Text
    if (-not $completedBeforeDeadline) {
        throw "Codex did not complete the Phase 2 probe within $TimeoutSeconds seconds."
    }
    if ($process.ExitCode -ne 0) {
        $diagnostics = @(
            $stderr -split "`r?`n" |
                Where-Object { $_ -match '(?i)outlook_classic|mcp|tool|error' } |
                Select-Object -Last 12 |
                ForEach-Object { Protect-DiagnosticText $_ }
        ) -join ' | '
        throw "Codex exited with code $($process.ExitCode) during the Phase 2 probe. Diagnostics: $diagnostics"
    }

    $events = [System.Collections.Generic.List[object]]::new()
    $lineNumber = 0
    foreach ($line in @($stdout -split "`r?`n")) {
        $lineNumber++
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }
        try {
            $event = $line | ConvertFrom-Json -ErrorAction Stop
        }
        catch {
            throw "Codex emitted malformed JSONL at line $lineNumber."
        }
        if ($null -eq $event -or (Get-OptionalProperty -InputObject $event -Name 'type') -isnot [string]) {
            throw "Codex emitted a JSONL value without an event type at line $lineNumber."
        }
        $events.Add($event)
    }
    if ($events.Count -eq 0) {
        throw 'Codex emitted no JSONL events.'
    }

    $allowedEventTypes = @(
        'thread.started',
        'turn.started',
        'item.started',
        'item.updated',
        'item.completed',
        'turn.completed'
    )
    foreach ($event in $events) {
        if ($event.type -cnotin $allowedEventTypes) {
            throw 'Codex emitted an unexpected or unsuccessful event type.'
        }
    }
    $turnCompletions = @($events | Where-Object type -eq 'turn.completed')
    if ($turnCompletions.Count -ne 1 -or $events[$events.Count - 1].type -ne 'turn.completed') {
        throw 'Codex did not emit exactly one successful terminal turn.completed event.'
    }
    if (@($events | Where-Object type -eq 'thread.started').Count -ne 1 -or
        @($events | Where-Object type -eq 'turn.started').Count -ne 1) {
        throw 'Codex did not emit exactly one thread and one turn start event.'
    }

    $allowedItemTypes = @('reasoning', 'agent_message', 'mcp_tool_call', 'todo_list')
    $mcpEvents = [System.Collections.Generic.List[object]]::new()
    $completedMcpEvents = [System.Collections.Generic.List[object]]::new()
    $completedAgentMessages = [System.Collections.Generic.List[object]]::new()
    $mcpItemIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($event in $events) {
        if ($event.type -notlike 'item.*') {
            continue
        }
        $item = Get-OptionalProperty -InputObject $event -Name 'item'
        $itemType = Get-OptionalProperty -InputObject $item -Name 'type'
        if ($null -eq $item -or $itemType -isnot [string]) {
            throw 'Codex emitted an item without a valid type.'
        }
        if ($itemType -cnotin $allowedItemTypes) {
            $safeItemType = Protect-DiagnosticText $itemType
            if ($itemType -eq 'error') {
                $safeErrorMessage = Protect-DiagnosticText (Get-OptionalProperty -InputObject $item -Name 'message')
                throw "Codex emitted disallowed item type '$safeItemType': $safeErrorMessage"
            }
            throw "Codex emitted disallowed item type '$safeItemType'."
        }
        if ($itemType -eq 'mcp_tool_call') {
            if ($item.server -cne 'outlook_classic' -or $item.tool -cne 'outlook_status') {
                throw 'Codex attempted an MCP tool other than outlook_classic outlook_status.'
            }
            $itemId = Get-OptionalProperty -InputObject $item -Name 'id'
            if ($itemId -isnot [string] -or [string]::IsNullOrWhiteSpace($itemId)) {
                throw 'The Codex MCP tool event did not contain an item identifier.'
            }
            $null = $mcpItemIds.Add($itemId)
            $mcpEvents.Add($event)
            if ($event.type -eq 'item.completed') {
                $completedMcpEvents.Add($event)
            }
        }
        elseif ($itemType -eq 'agent_message' -and $event.type -eq 'item.completed') {
            $completedAgentMessages.Add($item)
        }
    }
    if ($mcpItemIds.Count -ne 1 -or $mcpEvents.Count -eq 0 -or $completedMcpEvents.Count -ne 1) {
        throw 'Codex did not perform exactly one MCP tool call.'
    }

    $toolCall = $completedMcpEvents[0].item
    if ($toolCall.status -cne 'completed' -or
        $null -ne (Get-OptionalProperty -InputObject $toolCall -Name 'error')) {
        throw 'Codex did not complete the expected outlook_classic outlook_status call.'
    }

    $toolResult = Get-OptionalProperty -InputObject $toolCall -Name 'result'
    if ($null -eq $toolResult -or $toolResult -is [string]) {
        throw 'The Codex MCP event did not contain a native tool result object.'
    }
    $contentProperty = Get-PropertyEntry -InputObject $toolResult -Names @('content')
    if ($null -eq $contentProperty -or $contentProperty.Value -isnot [System.Array]) {
        throw 'The Codex MCP event did not contain the native content array.'
    }
    $structuredProperty = Get-PropertyEntry `
        -InputObject $toolResult `
        -Names @('structured_content', 'structuredContent')
    if ($null -eq $structuredProperty -or
        $null -eq $structuredProperty.Value -or
        $structuredProperty.Value -is [string]) {
        throw 'The Codex MCP result did not contain native structured content.'
    }
    $structuredContent = $structuredProperty.Value

    Assert-ExactProperties `
        -InputObject $structuredContent `
        -Expected @('ok', 'operationId', 'data', 'partial', 'warnings') `
        -Context 'The outlook_status success envelope'
    if ($structuredContent.ok -isnot [bool] -or $structuredContent.ok -ne $true -or
        $structuredContent.operationId -isnot [string] -or
        $structuredContent.operationId -cnotmatch '^[0-9a-f]{32}$' -or
        $structuredContent.partial -isnot [bool] -or $structuredContent.partial -ne $false) {
        throw 'The outlook_status success envelope flags or operation ID were invalid.'
    }
    $warningsProperty = $structuredContent.PSObject.Properties['warnings']
    if ($null -eq $warningsProperty -or
        $warningsProperty.Value -isnot [System.Array] -or
        @($warningsProperty.Value).Count -ne 0) {
        throw 'The outlook_status success envelope did not contain an empty warnings array.'
    }

    $data = $structuredContent.data
    if ($null -eq $data -or $data -is [string]) {
        throw 'The outlook_status success envelope did not contain a data object.'
    }
    Assert-ExactProperties `
        -InputObject $data `
        -Expected @('hostState', 'listenerReady', 'version') `
        -Context 'The outlook_status data object'
    if ($data.hostState -isnot [string] -or
        $data.hostState -cne 'online' -or
        $data.hostState.Length -gt 32 -or
        $data.hostState -match '[\x00-\x1F\x7F]' -or
        $data.listenerReady -isnot [bool] -or
        $data.listenerReady -ne $true -or
        $data.version -isnot [string] -or
        [string]::IsNullOrWhiteSpace($data.version) -or
        $data.version.Length -gt 64 -or
        $data.version -match '[\x00-\x1F\x7F]') {
        throw 'The Codex tool result did not contain the expected bounded online ready status.'
    }

    if ($completedAgentMessages.Count -ne 1 -or
        $completedAgentMessages[0].text -isnot [string] -or
        $completedAgentMessages[0].text.Trim() -cne 'MCP_PHASE2_CODEX_OK') {
        throw 'Codex did not return exactly one expected Phase 2 completion marker.'
    }

    $codexVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo($codexCommand.Source).ProductVersion
    $codexIdentity = if ([string]::IsNullOrWhiteSpace($codexVersion) -or $codexVersion -eq '0.0.0.0') {
        [IO.Path]::GetFileName($codexCommand.Source)
    }
    else {
        "codex-cli $codexVersion"
    }

    [pscustomobject]@{
        CodexCli = Protect-DiagnosticText $codexIdentity
        Server = $toolCall.server
        Tool = $toolCall.tool
        ToolCallStatus = $toolCall.status
        HostState = $data.hostState
        ListenerReady = $data.listenerReady
        ResultValidated = $true
        ProcessTreeClosed = $processTreeClosed
    }
}
finally {
    $cleanupFailures = [System.Collections.Generic.List[string]]::new()
    if ($null -ne $processJob) {
        try {
            $processJob.TerminateAndVerify(5000)
        }
        catch {
            $cleanupFailures.Add('The Codex process-tree job could not be emptied and verified.')
        }
        try {
            $processJob.Dispose()
            $processJob = $null
        }
        catch {
            $cleanupFailures.Add('The Codex process-tree job could not be closed.')
        }
    }
    if ($null -ne $process) {
        try {
            if (-not $process.WaitForExit(5000)) {
                $process.Kill()
                if (-not $process.WaitForExit(5000)) {
                    $cleanupFailures.Add('The gated Codex launcher did not stop within five seconds.')
                }
            }
        }
        catch {
            $cleanupFailures.Add('The Codex process exit could not be verified.')
        }
    }
    if ($null -ne $stdoutCapture -and $null -ne $stderrCapture) {
        try {
            Wait-ForOutputTasks -StandardOutput $stdoutCapture -StandardError $stderrCapture
        }
        catch {
            $cleanupFailures.Add($_.Exception.Message)
        }
    }
    if ($null -ne $process) {
        $process.Dispose()
    }
    if ($null -ne $launchGate) {
        $launchGate.Dispose()
    }
    [Environment]::SetEnvironmentVariable($tokenVariable, $previousToken, 'Process')
    if ($cleanupFailures.Count -gt 0) {
        throw "Codex probe cleanup failed: $($cleanupFailures -join ' ')"
    }
}
