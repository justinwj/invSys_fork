# Slice 4be-A: Users recording storage failures

Last verified2026-10-12 UTC. GREEN420/420 qualifies this test extension;4be-A remains incomplete.
Candidate `deploy/validation-guide-outcome-choices-01` remains unpromoted.

This extends evidence for the existing v4.11 D18 Users write contract: optional
recording must neither prevent nor repeat a business write, must show one fixed
tracking notice on failure, and must leave honest incomplete evidence. No runtime
implementation, outcome meaning, authority rule or recording eligibility changes.

`Slice4beAdminUsersRecordingStorage.ps1` extends the packaged Users gate after its
prior353 checks. For each Create/Update, Deactivate and Delete handler, it redirects
the attempt or SAVED record's temporary path to a nonexistent child immediately
before the real file Open. The real error handler records that Open failed;
the command/save path is observed without replacing its execution or persistence.

Each case requires one Core invocation and one persistence call, the exact fresh
saved business state read from Auth in read-only mode, one fixed tracking notice,
no leaked owned workbook or pending file, unchanged other Auth/inventory and
immutable earlier activity. Failed attempts leave no invented record; failed
outcomes retain only REQUESTED, with no private payload, source reference or false
conclusion. Only generated fixtures are mutated; they are restored byte-for-byte.

## Evidence

Paths below are relative to ignored `reports/runtime/`. Runtime data is not committed.

- Initial probe-installation attempt: controller
  `auth-read-controller/60929acafb014bbbbe56019dd3c0b791`; no behavioral tests ran.
  It restored settings/packages/tooling and closed Excel. Not product RED.
  Statement matching now uses exact trimmed lines, case-insensitively, to tolerate
  compiled VBA casing and preserve original statements.
- Expanded gate:420/420, retaining all353 prior check names in order. Controller
  `auth-read-controller/ea7079a2cb934673bfabd5e33738132d`; worker
  `slice4be-auth-read/fc59ae3771124fb6aebb4035509f8ec0`. Five instrumented compiles,
  six real Open failures, exact log/JSON, settings/package/tooling preservation and
  normal closure pass. Clean delayed native audit through20:32:31.9603264UTC.
- `admin-users-storage-static-01`:342 components/6428 procedures/140068 lines,
  all28 caps, dynamic9/45 and196 duplicates unchanged;505 parsers/three schemas pass.

The product is unchanged, so this is test-first coverage of existing behavior;
no artificial RED or implementation fix is claimed. All five package hashes match
the qualified editor build. Reuse that candidate's cold compiles and chain32/
live-role48/Create15 evidence; the Users form's prior visible evidence remains
unchanged. No native warning-dialog acceptance is claimed: the existing fixed
notice/confirmation seams observe messages while preserving real handlers.
See the Users recording and expected-outcome choices receipts for exact scopes.

The subsequent context-entry correction is tracked separately in
`plan022_slice4be_admin_users_context_results.md`. Other context boundaries,
continuous recordings, exact publication, independent
saved-guide conclusions, other controls and native/layout acceptance remain open.
D18-LOCAL-01 and D16-VBA-LOAD-01 remain unapproved.
