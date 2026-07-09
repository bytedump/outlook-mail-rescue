# Manual E2E — Version A (auto-detect identity)

GitHub runners have no classic Outlook, so the COM/GUI paths cannot be unit-tested. This
file is the explicit record of **what is verified by CI** vs **what must be validated by
hand against a live profile**, and **what wiring is still pending**.

> ⏳ **Time-critical:** the dev box's classic Outlook is on an **expiring grace license**.
> Step 1 (the probe) is the core hypothesis of Version A and the only license-dependent
> check — do it first, while the window is open.

Profile under test: `owner@company.com` (work M365), Windows login `user`.

---

## What CI already proves (no Outlook needed)

`lint + Pester (228) + gitleaks`, green on the PR. Every guard's **logic** is unit-tested
through pure helpers; only the live COM/GUI glue below is unproven.

| Helper (unit-tested) | Guard |
|---|---|
| `Resolve-IdentityToken`, `Get-IdentityPstFileName` | #1, #10, filename |
| `Resolve-DetectedIdentity` | #11 tier |
| `Get-PrimarySmtpFromEntry` | #8 SMTP read fallback order |
| `Invoke-WithRpcRetry` | #9 logon retry |
| `Resolve-OwnedProcessId`, `Test-ProcessExited` | #1 ownership |
| `Test-CanAutoFill` | #7 prefill |
| `Test-ValidExportName` | #6 export guard |
| `Test-PathUnderOneDrive` | #12 OneDrive root |
| `Get-ExportOutcome` | #13 degraded |
| `Compare-ReadbackIdentity` | #4 re-read |

---

## Step 1 — validate the COM probe (LIVE, do this first)

`Get-DetectedOwnerIdentity` (in `src/OutlookDetect.ps1`) composes the tested helpers over a
real Outlook COM session. It is additive (nothing in the running app calls it yet), so it
cannot have regressed anything. Validate it standalone — **run interactively** (so a
profile picker / license nag can't hang a non-interactive shell):

```powershell
# Windows PowerShell 5.1, matching Outlook's bitness (x64 here). Outlook should be open.
cd C:\Users\user\githubProjects\outlook-mail-rescue-version-a
powershell.exe -NoProfile -Command ". .\Invoke-MailRescue.ps1 -LoadOnly; Get-DetectedOwnerIdentity | Format-List *"
```

**Expected:**
```
Tier           : smtp
PrimarySmtp    : owner@company.com
DisplayName    : <the display name on the profile>
NameSourceText : owner@company.com
HasIdentity    : True
```

Pass criteria:
- `Tier = smtp` and `PrimarySmtp` is the full address (proves the `0x39FE001F` read, guard #8).
- No profile-picker dialog appeared (proves `Logon(showDialog:$false)`, guard #3).
- After it returns, **no new lingering `OUTLOOK.EXE`** if Outlook was already running
  (ownership/Quit, guard #1) — check Task Manager. If Outlook was NOT running and the
  probe started it, exactly that instance should have quit.

If `Tier` comes back `none`/`displayname`/`exchange`, capture the output and the live
`Namespace.CurrentUser.AddressEntry.Type` so the fallback order can be tuned.

> ✅ **DONE (probe validated live, 2026-06-29):** `Tier=smtp`,
> `PrimarySmtp=owner@company.com`, no profile picker, owned Outlook quit.

---

## Step 2 — validate the prefill wiring (LIVE, commit `af8fe2e`, NOT yet validated)

The probe is now wired into the GUI (background runspace -> timer -> `Test-CanAutoFill`
prefill + "Detect owner" button). Parse/analyzer/Pester are green but **WinForms behavior
is unproven** (construction is not unit-tested). Launch the app and check both ownership
paths:

```powershell
cd C:\Users\user\githubProjects\outlook-mail-rescue-version-a
powershell.exe -NoProfile -File .\Invoke-MailRescue.ps1
```

- **Path A (attach):** open classic Outlook first, then launch the app. The Mailbox field
  should auto-fill `owner@company.com` on show, footer `Detected owner:
  ... [smtp]`, and **Outlook must stay open** (attach, not own).
- **Path B (button):** Outlook closed, launch the app (field empty, no auto-launch). Click
  **Detect owner** -> footer "Detecting...", field fills, and the Outlook the probe started
  **quits by itself**.

Pass -> the increment is confirmed. Fail -> fix before building further.

> **Result (2026-06-29) — both paths PASS:**
> - **Path B (Detect owner button) = PASS** — fills the correct identity, footer shows the
>   detected owner + tier; clicking always works (the button bypasses the startup gate).
> - **Path A (auto-prefill on show) = PASS** — with the **classic** `OUTLOOK.EXE` running at
>   launch the Mailbox field auto-fills on show, no click needed. The first run that came up
>   empty had the **new** Outlook (`olk.exe`) open, not classic, so `ClassicRunning` was
>   correctly `false` — not a bug. Confirmed by re-testing with classic verified running
>   (`Get-Process -Name OUTLOOK,olk` showed `OUTLOOK` only) → field auto-filled.
>
> By design the startup auto-probe is gated on `$info.ClassicRunning` (`Gui.ps1` Add_Shown
> ~:620; `ClassicRunning` = `Get-Process -Name 'OUTLOOK'`, `OutlookDetect.ps1` ~:314) so the
> app never launches Outlook unprompted; the new Outlook (no COM) correctly falls through to
> the manual button. **#5/#7 confirmed, no debt.**

---

## Step 3 — auto-launch classic Outlook, KEEP-OPEN (#7, LIVE)

Overrides the old attach-only gate: at startup, if classic Outlook is installed but not
running, launch it (`Start-OwnedOutlook`) and keep it open; the app owns that PID and quits
it on close (`Stop-OwnedOutlook`). If classic is already running, attach only. **No export
needed.**

```powershell
cd C:\Users\user\githubProjects\outlook-mail-rescue-version-a
powershell.exe -NoProfile -File .\Invoke-MailRescue.ps1
```

- **Cold (Outlook closed):** app launches classic, it stays open, field auto-fills. Watch
  for the license/grace nag on the cold launch. Close the app → the app-launched Outlook quits.
- **Warm (classic already open):** app must NOT launch a second and must NOT quit it on close.

> ✅ **DONE (validated live 2026-06-30):** cold launch opens + keeps classic and quits it on
> close; warm attach leaves the user's Outlook untouched. Also fixed a latent guard-#1 bug this
> surfaced: `.Quit()` closes the *connected* Outlook (not a PID), so a transient COM-activation
> PID made the probe/export quit an already-open session — now we quit only when no `OUTLOOK.EXE`
> existed before we launched (`$before` empty), in `Get-DetectedOwnerIdentity` and
> `Export-MailboxToPst`.

---

## Step 4 — export-path batch: #4b + #3 + #13 (LIVE, one full export)

These touch the export/result path, so they validate together in a single real export
(~24 min). Use a **fresh output folder** (e.g. `C:\Exports\run2`) so the 2.81 GB backup at
`C:\Exports\owner@company.com.pst` is untouched and anti-clobber does not block.

> **Scope note (2026-06-30):** the PST is written **directly to its final `<email>.pst` name**
> — no `.partial`/rename and no PST root-folder rename. `AddStoreEx` only accepts a `.pst`
> extension (a `.partial` is rejected as "not a valid Outlook data file"), and the owner chose
> to keep the standard name with no in-PST renaming. #3 is therefore just the **FormClosing
> warning** (block close mid-export). **#14 (JSON manifest) was validated then removed at
> the owner's request — unwanted.**

```powershell
cd C:\Users\user\githubProjects\outlook-mail-rescue-version-a
powershell.exe -NoProfile -File .\Invoke-MailRescue.ps1
```

- **#4b re-read** (`Export-MailboxToPst -ExpectedSmtp`): matching identity must NOT false-abort
  → still `28122/28122`. Enforced only when the field equals the auto-detected SMTP; a manual
  name passes empty `-ExpectedSmtp` and skips the check. Mismatch aborts BEFORE `New-PstStore`
  (covered by `Compare-ReadbackIdentity` unit tests; hard to force live).
- **#3 close-warning** (FormClosing guard): closing the window while an export is active is
  **blocked** with "Export in progress. Wait for it to finish before closing." No file renaming
  (PST keeps its standard name; detaching the PST inside Outlook itself is not interceptable by
  the app, only the app window close is guarded).
- **#13 degraded title**: result dialog reads "Export incomplete" + reasons only when
  `CopiedItems < SourceItems` (items missing) or a stall/block occurred; an equal *or larger*
  copy is "Export complete" (`Get-ExportOutcome` wired into the return as `.Outcome`).
  **Revised 2026-07-09:** `CopiedItems > SourceItems` used to be flagged "possible duplication";
  it is now an informational note, since `CopyTo` copies each subtree once and the extra items
  are just mail that arrived mid-export. Pure logic, unit-tested — no live re-validation needed.

Pass criteria: full export `28122/28122`, mid-export close blocked with the warning, title
"Export complete".

> ✅ **DONE (validated live 2026-06-30):** full export to the final `<email>.pst` name; #4b did
> not false-abort; closing the app mid-export was blocked with the warning; result title
> "Export complete". The `.json` manifest (#14) was validated live too but then **removed at
> the owner's request**.
> **Remaining follow-up:** the elevated PST/OST scan opens an oversized PowerShell console
> window → to be hidden. Tracked separately.

---

## Pending wiring (NOT done — needs live feedback, touches validated code)

> #5 (Detect button / attach-only) and #7 (prefill) are WIRED (commit `af8fe2e`) and await
> Step-2 live validation. The rest below is not started.

The pure logic is ready, but **none of it is wired into the running GUI/export yet**. This
glue lives in the WinForms event handlers and the validated `Export-MailboxToPst`; it can
only be confirmed by running the app, and a blind rewrite risks the class of bug that the
`Application.Run`/`$timer` scope issue already caused once. Do each incrementally, verify
live, commit:

- **#5 attach-only probe + "Detect owner" button** (`Gui.ps1`): call `Get-OutlookInfo`; if
  `ClassicRunning`, auto-probe via `Get-DetectedOwnerIdentity`; else show a "Detect owner"
  button. No process-launching COM at startup.
- **#7 prefill** (`Gui.ps1` `TextChanged`): on a probe result, set the field via
  `Test-CanAutoFill` (track last auto-set value; upgrade in place; never clobber typed text).
- **#6a export gate + filename** (`Gui.ps1`) — **DONE, validated live (2026-06-29)**: Export
  button starts disabled and is gated on `Test-ValidExportName` via `Update-ExportEnabled`
  (fires on `TextChanged` + `Add_Shown`); empty / `user` / cue text keep it disabled, a valid
  identity enables it. PST name built with `Get-IdentityPstFileName` (`$null` -> refuse, no
  `user.pst`); confirmed name string `owner@company.com.pst`. Full-export
  re-confirm of the new name deferred to #1.
- **#6b cue banner** (`Gui.ps1`) — **DONE, validated live (2026-06-29)**: `EM_SETCUEBANNER`
  (P/Invoke `SendMessage`) shows the grey `owner@company.com` placeholder in the empty
  Mailbox field; wrapped in try/catch so it can never break the GUI. **5.1-vs-7 gotcha
  caught live:** `Add-Type -MemberDefinition` already injects `using
  System.Runtime.InteropServices;`, so `-UsingNamespace` duplicated it and failed only under
  the PS 5.1 csc warning-as-error compiler (pwsh7/Roslyn deduped, so Pester was green) —
  fixed by dropping `-UsingNamespace`. Confirms the rule: lint+Pester run in pwsh7, the app
  runs in 5.1; native interop must be smoke-tested under 5.1.
- **#12 OneDrive output** (`Gui.ps1`) — **DONE, validated live (2026-06-29)**: `Get-OneDriveRoots`
  gathers roots (env `$env:OneDrive*` + HKCU `Accounts\*\UserFolder`); the export click warns
  via a YesNo MessageBox when `Test-PathUnderOneDrive` is true and aborts on No. Collector
  proven under PS 5.1 (returns `C:\Users\user\OneDrive`; sub=True, `C:\Exports`=False). Live:
  OneDrive folder -> warning -> No aborts clean.
- **#4a confirm dialog** (`Gui.ps1`) — **DONE, validated live (2026-06-29)**: the export click
  shows a YesNo "Confirm export" dialog echoing Owner (detected DisplayName, or "(entered
  manually)" when the field no longer matches the detected SMTP) + Mailbox + `$env:COMPUTERNAME`
  + the resolved `Save to` path; No aborts before anything runs. The probe result is retained
  in `$script:MrProbe.LastDetected` for the owner line. Validated both auto and manual paths.
- **#4b re-read in runspace** — pending, **bundled into #1**: re-read `CurrentUser` inside the
  export runspace once Outlook is logged on and branch on `Compare-ReadbackIdentity`
  (`mismatch` -> abort, `unverified` -> warn). Touches `Export-MailboxToPst`, so it rides the
  single full-export re-confirm with #1/#3.
- **#1 ownership rewire** (`ComExport.ps1`) — **DONE, validated live (2026-06-29)**: replaced
  `$startedByUs` with the PID-snapshot + `Resolve-OwnedProcessId` (quit only an instance we
  started) + a `Test-ProcessExited` exit poll. Full export re-confirmed: `items 28122/28122`,
  stores=1, folders=22 (count grew from 27341 as the live mailbox grew; source==copied means
  nothing skipped), clean name `owner@company.com.pst` (2.81 GB, unlocked),
  and no lingering `OUTLOOK.EXE` afterwards. The first run logged a false "still running ~5s
  after Quit" WARN because a large export's Outlook takes >5s to flush/close; the poll was
  widened to ~30s (timing constant only, logic unchanged, not re-timed — low risk).
  - **Interruption safety (found live 2026-06-29):** there is no Cancel during export (a
    deliberate choice — cancelling a COM `CopyTo` mid-flight is riskier than waiting), but
    closing the window mid-export is an uncontrolled kill. With the new clean `<email>.pst`
    name a partial would masquerade as a complete backup AND block re-export via anti-clobber.
    Fold two fixes in here: (a) a `FormClosing` guard that blocks/warns while an export is
    active; (b) write to a temp/partial name and rename to the final `<email>.pst` only on
    success, so a clean-named PST always means complete. (b) touches the validated export —
    confirm with the owner before changing it. #13 supplies the "Export incomplete" wording.
- **#13 degraded title** (`ComExport.ps1` ~:283, `Gui.ps1` ~:446): title the result dialog
  with `Get-ExportOutcome().Title` and list `.Reasons` instead of always "Export complete".

> **Note:** #4b / #3 / #13 are DONE and validated live (see Step 4). #14 (JSON manifest) was
> built and validated, then removed at the owner's request. Only the oversized scan-console
> window remains as a follow-up.

## Regression check (must stay intact)

- Full COM export still copies **27341/27341** items to a Unicode PST, reopen + recount matches.
- Bitness self-relaunch, elevated-scan JSON IPC, new↔classic toggle all still work.

---

## Decisions taken autonomously (veto anytime)

- **PST filename = full primary SMTP, literal, no timestamp** -> `owner@company.com.pst`.
  Safe because the existing anti-clobber (`Test-Path` + `New-PstStore` throw) blocks an
  overwrite instead of silently replacing a prior export.
- **Exchange tier file name = DisplayName-preferred** (legacyDN only as last resort).
- **Degraded title = "Export incomplete"** (strongest contrast with "Export complete").
- **Prefill = upgrade-idempotent**: a better detection overwrites while the field still
  holds our last auto-set value; a user edit ends auto-fill for good.
