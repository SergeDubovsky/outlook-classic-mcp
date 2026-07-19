# ADR 0001: Event-driven Outlook STA wake-up

Status: Accepted
Date: 2026-07-19

## Context

The implementation plan initially selected `Control.BeginInvoke` on a hidden WinForms control as the Outlook STA wake-up mechanism. During repeated Release-mode Outlook start/close testing, cross-thread `BeginInvoke` calls could return successfully without the queued dispatcher callback running. A UI-owned WinForms timer proved the rest of the dispatcher and lifecycle design, but a continuously running timer would wake idle Outlook unnecessarily.

## Decision

Create a hidden WinForms control and force its window handle on the captured Outlook startup STA. Keep delegates and completion sources in the bounded managed queue; worker threads receive no control or COM object. When work becomes available, post one private `WM_APP` message to the captured handle with `PostMessage`. The control drains one item per message on its owner STA and posts another wake-up only when work remains.

The dispatcher guards the whole drain against nested Windows message pumping, never uses `SendMessage`, and fails queued work closed if a wake-up cannot be posted. Shutdown stops admission, completes queued-but-unstarted work, allows active work to finish, unpublishes the handle under the dispatcher lock, removes an outstanding private message, and disposes the control on its owner STA.

## Consequences

- Idle Outlook receives no periodic dispatcher wake-ups.
- Outlook work remains serialized on the captured UI STA.
- The private window message is only a wake-up signal; request data never crosses the native message boundary.
- Handle recreation is unsupported. An unexpected handle loss makes the dispatcher unavailable rather than moving work to another thread.
- Unit tests cover nested message pumping, capacity, cancellation, idempotent shutdown, and terminal completion of accepted work. The Release smoke suite verifies the same-STA path across three separate Outlook processes.
