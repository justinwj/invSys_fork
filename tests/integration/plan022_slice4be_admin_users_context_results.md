# Slice 4be-A: Users recording context entry

Last verified2026-10-12 UTC. RED254/3 -> GREEN493/493 and fresh chain qualification pass.
4be-A remains incomplete. Unpromoted candidate:
`deploy/validation-admin-users-context-01`.

Under existing v4.11 D18/D8-USERS-01, changed live context must refuse a captured
Users write, avoid stale/new-context attribution, and show the fixed tracking
notice. The observer previously read live context before optional recording;
sign-out inside that call returned no action/notice, leaving its comparison stale.
`modAdminUsersAction.BeginObservation` now reads live context after the call.
The notice wording, captured inputs, Core authority and policy are unchanged.

`Slice4beAdminUsersRecordingContext.ps1` exercises actual Save/Deactivate/Delete
handlers with sign-out, actor switch and warehouse switch at recording entry.
It verifies the context change happened, one Core call with zero persistence,
the existing authority refusal, one fixed tracking notice, unchanged authorities/
inventory/earlier activity, no stale or later-context records and no leaked workbook.
Fixtures are Admin-generated; the probe never supplies a business result.

## Evidence

Runtime paths below are relative to ignored `reports/runtime/`.

- Focused RED on `validation-guide-outcome-choices-01`:254PASS/3FAIL, solely
  `UsersContext.SignOut.<Save|Deactivate|Delete>.OneFixedTrackingNotice`.
  Controller `auth-read-controller/d99e20b0cdd240da91cc5d096b8237ef`; worker
  `slice4be-auth-read/63e3e9a501014c7caf30f13e6183a761`.
  All nine context changes preserve refusal/authority/attribution; retained
  recording and six real storage-failure cases pass. Five instrumented compiles,
  exact log/JSON, preservation and normal closure pass. Native audit clean through
  20:45:20.2440900UTC. This focused run is not the full420-check authority gate.
- Expanded GREEN493/493 retains all420 prior checks and all257 focused RED checks
  in order. Controller `auth-read-controller/87af6ea5187149d8a20cd3a9c12dc0f8`;
  worker `slice4be-auth-read/59c39d36b2234456993871d51027953d`. Nine context changes
  and six real storage failures pass. Five instrumented compiles, exact log/JSON,
  settings/package/tooling preservation and normal closure pass. Native audit clean
  through21:04:23.1115016UTC. Four reviewed captures retain healthy/invalid Auth,
  the full context refusal and Settings Save, with credential fields blank.
  Existing long-path/account-copy-button clipping remains open.
- `admin-users-context-build-01`: five cold compiles and Operations cold start;
  334 compiled components unchanged, only Admin/modAdminUsersAction changed.
  Frozen predecessor/settings preserved; clean native audit through20:46:30.3123714UTC.
- `admin-users-context-static-01`:342 components/6428 procedures/140068 lines;
  all28 size caps, dynamic9/45 and196 duplicate groups unchanged;506 parsers and
  all three schemas pass.
- Fresh chain32/live-role48/Create Warehouse15 retains the prior ordered checks.
  Controller `admin-users-context-regression/chain-b8a65ba3c8a24f879238c10db58322fc`,
  21:04:34.2835261-21:10:25.1116273UTC; unassisted closure, settings/package/tracked-
  report preservation and clean delayed native audit through21:10:37.1116273UTC.
  Three cleanup waits completed normally, without intervention.

Confirmation and warning observations use the existing explicit test seams. This
does not establish ordinary native warning presentation or broader form/layout
acceptance. Context changes at other boundaries, continuous recordings, exact
publication, independent saved-guide conclusions and other controls remain open.
No local-recording or replay eligibility decision is made.
