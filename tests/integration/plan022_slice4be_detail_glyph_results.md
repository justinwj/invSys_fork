# 4be-A Event Detail glyph rendering

Last verified: 2026-10-12 UTC. This corrects v4.11 D18's existing complete
multiline-text requirement; no new contract, recording ID or permission.
Candidate: `deploy/validation-detail-glyph-01`, unpromoted. Full4be-A remains open.

## Findings and D13

Refreshing the older multiline proof on directory-01 initially passes87/88.
The native list-scroll input lands at the right-arrow boundary at96DPI and moves
only a little. Moving the test point one scrollbar width inside the track gives
88/88 with the same detector/threshold. Its existing calibration remains8/8.
This is a harness correction, not product RED; preserve the failed attempt.
Current-candidate Viewer separately passes98/98 with three reviewed captures.
Receipt: `reports/runtime/directory01-viewer-refresh-verification.json`.

Image review rejects two final-line captures despite88/88: the label clips the
bottom pixel row of the final `END`. Temporary height enlargement of the actual
disposable form reveals366 additional text pixels. The new packaged test restores
the original label before subsequent checks and compares only the frame interior,
excluding scrollbar/border/input pixels. Pure detector calibration9/9 at100/150/200%
rejects unchanged images and unrelated outside changes.

- Focused product **RED89/1**: `EventDetail.Multiline.FinalGlyphsUnclipped` only.
  Runtime root `slice4be-viewer-detail/ced46d8afc2140fea0550769e02b41b3`.
- Product fix: `frmEventDetail.AddMultilineLabel` retains measured width, disables
  AutoSize after measurement and adds two points of height before computing extents.
  Original text, profile order, read-only behavior and owner/context remain.
- Focused **GREEN90/90**, retaining all88 prior checks in order. The same temporary
  enlargement now reveals **zero** extra pixels. Runtime root
  `slice4be-viewer-detail/dafcc64bda4c4f0299701b52cac563ff`.
- Fourteen principal images and two explicitly diagnostic oracle images reviewed.
  Complete final glyphs, multiple-field order, scrolling, resize/restore and layout
  are visible. Diagnostic enlargement is not presented as product behavior.

Both focused runs pass five instrumented compiles, close normally, restore settings
and preserve package/tooling hashes and generated authority/unknown-column bytes.
Delayed Application audit through16:17:21.1997476UTC has zero failures.
Receipt: `reports/runtime/detail-glyph-focused-verification.json`.

## Qualification and limits

Five cold compiles and Operations cold start pass. Compiled comparison isolates
the change to Operations/frmEventDetail;329 other components are unchanged.
Build/settings/frozen directory-01 preservation and delayed native audit pass:
`reports/runtime/detail-glyph-build-01/verification.json`.

Rebuilt-candidate Viewer **98/98** and packaged smoke **86/86** retain the exact
preceding check orders. Three Viewer captures are reviewed. Both smoke hosts exit
unassisted. The diagnostic smoke copy replaces forced termination with a failure
and guards deletion of its resolved generated scratch folder; business assertions
are unchanged. Settings, tracked reports and package pins are preserved.
Combined delayed audit through16:19:49.8680565UTC has zero Application failures.
Receipt: `reports/runtime/detail-glyph-regression-verification.json`; Viewer root
`slice4be-viewer-published-read/1c9bbb1248364e66b7aa196076f467a5`; smoke controller
`detail-glyph-regression/smoke-84c76cc159744c7c8a0a5c62336a2de9`.

Regenerated static01:337 components,6412 procedures,139747 lines (+2);28 oversized
caps unchanged; dynamic calls9/45 and duplicate groups194 unchanged. Three schemas
validate; all495 PowerShell files parse. Receipt:
`reports/runtime/detail-glyph-static-01/ratchet-verification.json`.

This completes the focused rendering correction and its applicable local gates.
No fresh full live-role/Release1 chain is claimed for glyph-01; that qualification
remains blocked independently by the unresolved ordinary-loading failure.
No station setting was adopted; `ForceVBALoadFromSource` remains absent.
D16-VBA-LOAD-01 and D18-LOCAL-01 remain proposed, not approved. Earlier ordinary
chain failures and temporary source-loading controls retain their exact scope in
[the directory receipt](plan022_slice4be_admin_directory_columns_results.md).
This rendering correction does not repair or waive that independent native gate.
