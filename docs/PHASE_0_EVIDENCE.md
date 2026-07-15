# Phase 0 evidence

Status: in progress
Last updated: 2026-07-15

## Decisions

| Area | Decision |
|---|---|
| Repository | `SergeDubovsky/outlook-classic-mcp`, public |
| Product namespace | `OutlookClassicMcp` |
| License | Apache-2.0 |
| Endpoint | Exact `http://127.0.0.1:8765/mcp/` |
| Initial validation | Any CPU add-in, x64 Classic Outlook |
| Inspector | Deferred while the development machine remains on Node 20 |
| Consequential tools | Send and soft-delete remain absent until implementation and approval gates pass |

## Toolchain preflight

The implementation preflight verified:

- Visual Studio Community 2026 18.8 with Microsoft 365 development, VSTO, managed desktop development, and .NET Framework 4.8 SDK/targeting components.
- Full MSBuild 18.8 and the VSTO OfficeTools targets.
- Discovery of the installed `VSTO_Outlook15AddIn_CS` template and its `Office15Wizard`.
- .NET Framework 4.8 reference assemblies, VSTO runtime 10.0.60910, and .NET SDK 10.0.302.
- Classic Outlook x64 16.0.17932.20842.
- Port 8765 free before the proof.
- A non-elevated `HttpListener` start/stop at the exact endpoint, followed by release of the port within three seconds.
- No HTTP.sys URL ACL is required on the development machine.

`tools/preflight.ps1` reproduces the non-mutating checks and listener lifecycle proof.

## Untouched VSTO template proof

The installed Visual Studio template was instantiated through `VisualStudio.DTE.18.0`; its Office project wizard generated the stock classic VSTO project. Before retargeting or adding dependencies, full MSBuild compiled the original .NET Framework 4.7.2 template and generated its VSTO deployment manifests using a temporary non-exportable code-signing certificate supplied only through command-line properties.

The proof returned exit code 0. The certificate and its CNG private-key container were removed afterward, no certificate file was written, and all stock source/project hashes were unchanged. Visual Studio generated a classic `.sln`; no SDK conversion or hand-authored VSTO project was used. The template was instantiated through the installed Visual Studio automation model rather than mouse-driven GUI interaction; the same installed wizard and stock template produced the project.

Generated-file baselines:

| File | SHA-256 |
|---|---|
| `ThisAddIn.Designer.cs` | `259419964B326D447D5194495265F93C4A08600D1760CA1724614051168DC436` |
| `ThisAddIn.Designer.xml` | `CE48FE49C83A3F7B16E4E72EDEA95A59C0143EC40B321BBE269C811ADFD3920B` |
| `Resources.Designer.cs` | `6FB6715D3EE01AA9263A2119109A9CB8499600192880C70D2CE3132433892855` |
| `Settings.Designer.cs` | `671815A8F70DB2530DC85758C6ECB036A09AFD2B3ACA5E17D221E915FF8133BB` |

After that proof, the project was moved under `src`, retargeted to .NET Framework 4.8, pinned to C# 9.0 through `Directory.Build.props`, and connected to the Office-free Core and Transport projects. The generated files still match these baselines.

## Build and automated tests

Locked restore succeeds for all five projects. Debug and Release/Any CPU builds succeed with full Visual Studio MSBuild and ephemeral manifest signing. The isolated build refuses to overwrite pre-existing same-name VSTO development state, invokes `VSTOClean` for every attempted configuration, and verifies that its Outlook registration, VSTO inclusion/metadata entries, certificate, CNG container, and key file are absent afterward.

Automated test results:

| Assembly target | Passed | Failed |
|---|---:|---:|
| Core Tests `net10.0` | 4 | 0 |
| Core Tests `net48` | 4 | 0 |
| Transport Tests `net48` | 4 | 0 |

The tests prove the Phase 0 tool catalog is empty, the Core and Transport assemblies have no Office/VSTO references, and the real exact-prefix listener starts and releases its namespace.

Public Windows CI run [29451627206](https://github.com/SergeDubovsky/outlook-classic-mcp/actions/runs/29451627206) completed successfully. The hosted runner verified its Visual Studio Office/VSTO components and Office 15 PIA assembly identities, installed and verified the pinned Microsoft VSTO runtime 10.0.60917 redistributable, built the signed Release solution with locked restore, passed all 12 tests, and passed the generated-file, dependency-license, and deterministic-repository checks.

## Codex configuration gate

Codex CLI 0.144.4 strictly parses the committed repository configuration, including the bearer-token environment variable, default write approvals, and per-tool prompt settings. `tools/configure-codex.ps1` treats that reviewed project file as immutable, rejects a same-name global registration, proves the server appears only from this trusted checkout, and manages only the current-user token. It does not rewrite TOML. Global registration and structure-preserving global removal are deferred until a TOML-aware edit path is implemented.

The current desktop permission profile does not prove an unavoidable prompt because an effective top-level `never` approval policy can auto-approve MCP tools. The plan's fail-closed rule therefore remains active: send and soft-delete are not implemented or exposed.

## Remaining Phase 0 gates

- Save work and close the currently running Outlook instance gracefully, then perform the one interactive F5 proof that reaches `ThisAddIn_Startup` and confirms the add-in is Active rather than Disabled.
- Node 22 installation remains intentionally deferred; Inspector/conformance is not an early-phase gate.
