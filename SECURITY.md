# Security policy

## Reporting a vulnerability

Do not disclose suspected vulnerabilities in a public issue. Use GitHub's private security-advisory reporting for this repository. Include the affected version, reproduction steps, impact, and any suggested mitigation. Do not include real mailbox content, bearer tokens, Outlook identifiers, attachment data, or local paths containing sensitive information.

## Supported versions

Until the first public release, only the current `main` branch is supported. Security fixes will be applied to the latest supported release after versioned releases begin.

## Security scope

This project is intended for one interactive Windows user. It authenticates a literal IPv4 loopback endpoint and relies on the active Outlook profile's existing permissions. It does not claim to protect against administrators, malware, code injected into Outlook or the MCP client, or physical access to an unlocked session.

Email bodies, headers, links, and attachments are untrusted input. A security report is especially useful when it demonstrates authentication bypass, non-loopback exposure, route or path confusion, prompt-driven privilege escalation, unintended tool exposure, cross-store identity confusion, arbitrary file access, unsafe attachment handling, duplicate mutation/send behavior, sensitive logging, or Outlook process destabilization.
