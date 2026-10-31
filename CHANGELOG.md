# Changelog

## 2.3.1 (2026-10-31)

### Removed
- The cutterhead design notes and their trigger words, and two API claims that were never tested
  here (`RecordMacro` / `EditMacro` through IDispatch, a Node.js `winax` `SelectByID2` limit).
- Third-party credits from LICENSE, README, the skill's sources and this changelog: everything in
  the skill now comes from this project's own builds and the SOLIDWORKS type library.

## 2.3.0 (2026-10-31)

Lessons from the SolidWorks Test Runs of swlib 2.1/2.2 and of the author-credit helpers
(14 of 14 helper checks, 4 of 4 credit files INTACT, one SOLIDWORKS crash).

### Added
- §11.1 Troubleshooting playbook: symptom, first check, likely cause and fix, grouped into
  connection and crashes, return codes, and calls that return None.
- *The RPC server is unavailable* (-2147023174) means SOLIDWORKS crashed or exited: stop, don't
  retry (§11.1, TROUBLESHOOTING).
- After a crash, check for more than one SLDWORKS process before attaching again (§11.1).
- Decode error return codes as bit masks against swconst (§11.1).
- §8: hidden data in a model with a hidden attribute (`DefineAttribute`, `CreateInstance5` option 1),
  with a verified snippet; write and verify in the same process (a second process was followed by a
  crash, cause not proven).
- §8: `IGet3rdPartyStorage` returns None from an external script; use an attribute or an NTFS
  alternate data stream instead.
- §8: STEP header and PDF author fields (`stamp_step`, `stamp_pdf`), now proven.
- Rule 1.15: Windows paths as raw strings, and a control-character check on generated text.
- §16: the credit helpers in `swlib.py` and `tools/verify_credit.py`.
- Docs: README (coverage, repo tree, troubleshooting), `examples/README.md`, the `swlib.py` header,
  `CLAUDE.md` and `docs/USAGE.md` describe the playbook and the credit helpers.

## 2.2.0 (2026-10-31)

Lessons from two more builds: a vernier caliper (parts, assembly, 7 toleranced drawings; 3.5 h,
about 2 h avoidable) and light-tight vent grilles added to an existing housing (9 full rebuilds to
test small changes).

### Changed
- "Probe" is now **SolidWorks Test Run** everywhere. `examples/00_probe_template.py` is now
  `examples/00_test_run_template.py`. Rule 1.14 is mandatory and covers speed (time a small test
  run before any sketch with more than 40 segments or any face/edge loop on a detailed part).
- §0 lists the failure causes of all three builds.

### Added
- Rules 1.15 (safe script patching, a dry mode before destructive steps, no stdlib-like script
  names), 1.16 (confirm the style with a render), 1.17 (fits, clearances and standard values
  before modelling) and 1.18 (editing an existing part: backup with MD5, modified time, session
  check, renamed dry-run copy, named and foldered feature groups).
- §7.4 plane through 3 points; §7.6 oblique prism cut; §7.10 missing stainless grades; §7.12 big
  sketches, engraving behind a `NO_MARKING` configuration, entity names, turned parts.
- §8: save over an earlier result by path; author properties. §8.1 renders (model-to-pixel
  mapping, crop, report sheets, PDFs). §8.2 drawings through the API (sheets, views, toleranced
  dimensions, GD&T symbols, datums, notes, surface finish, dry run). §8.3 drawing layout and a
  GD&T checklist, with an automatic PDF collision check.
- §9.1 measuring an existing part with `RayIntersections`. §10.3 mates that check their own
  direction. §11 visible progress, background runs, feature groups. §12 skin diff, sight-line test,
  drawing checks. §14.2 precision parts, §14.3 vents, patterns and print orientation.
- `swlib.py` 2.2: `Asm.mate(verify=)` (retries flip and alignment in the same run), save by path
  (`close_other`), `plane_3pts`, `name_entity` / `entity`, `folder`, `delete_group`, `ray_hits`,
  `backup` (MD5) and `log`. All passed a SolidWorks Test Run on SOLIDWORKS 2025 SP1.2
  (2026-10-31): the self-checking mate ended at −40 mm as asked, deleting a feature group returned
  the volume exactly, and ray hits are type 17 at entry and 33 at exit.

## 2.1.0 (2026-10-29)

Lessons from a torsion-spring build: one part with configurations, a flexible sub-assembly and
a test rig. 28 of about 50 runs failed or gave a wrong result. Each cause is now a rule.

### Added
- §0: a table of the failure causes and the rule that prevents each one.
- Rules 1.13–1.16: import `swlib` and use one binding mode; probe before building; edit scripts with
  an editor, not with shell heredocs; confirm an ambiguous interpretation with a render before
  building the assemblies.
- §3.0: binding rules, early- vs late-bound side by side (zero-argument members, ByRef outputs,
  `NOTHING`, object arrays, `FeatureByName`, indexed and read-only property puts with `com_put`).
- §7.6: the measured helix hand, start and dimensions, with a hand check. A merging feature must
  overlap the body or keep a gap, never only touch it. Combine (Add).
- §7.8: editing an existing equation in place.
- §7.10: appearance per configuration; reading the material back, with a fallback.
- §7.11: configurations (plan the parent/child tree, activate then suppress, stale handles,
  dependents, check every configuration before saving).
- §9: exact extents with `GetExtremePoint`.
- §10: components in a chosen configuration, limit angle mates through a 90° datum, `newest()` and
  `mate_errors()`, Make Subassembly Flexible through the API, rebuild order, and the drag test (§10.8).
- §11: closing only the documents the script created. §12: per-configuration, mate-error,
  drag-test and two-pass interference checks. §14: torsion spring formulas.
- New constants and signatures (Combine, configurations, suppression, solving, measuring).
- `swlib.py` 2.1: `NOTHING`, `dispatch_array`, `com_put`, `newest`, `extent`, plane flip retry,
  helix and hand check, circular sweep, combine, `set_global`, configuration helpers,
  per-configuration colour, material read-back, configured `Asm.add`, angle-limit mates, mate
  errors, `make_flexible`, two-pass interference, and closing only the documents the script created.
- `examples/00_probe_template.py`.
- Every `swlib` 2.1 part and assembly helper was run early-bound on SOLIDWORKS 2025 SP1.2
  (2026-10-30). Those probes added: `SwApp` waits for a SOLIDWORKS that is still starting
  (*The interface is unknown*), SaveAs3 error 2049 means the path is too long, and interference
  from a parent assembly doesn't look inside sub-assemblies.

### Fixed
- Rule 1.6: `FeatureByPositionReverse(0)` is the newest feature only in parts.
- §7.3: the sketch transform no longer needs `IMathUtility`, which throws when late-bound.
- §10.3: the mate helper now passes the angle arguments (positions 9–11, radians).
- §10.7: flexible sub-assemblies are set through the API, not "tell the user to right-click".

## 2.0.0 (2026-10-29)

First public version, based on real product builds.

### Added
- Correct values for commonly mistyped constants: `swEndCondBlind` is `0`, `swEndCondThroughAll`
  is `1`, `swSelEDGES` is `1`, `swSelSKETCHES` is `9`; `swPlaneTop/Front/Right` don't exist, so
  planes are found by type order.
- Long argument lists work through IDispatch in pywin32 (23 and 27 arguments, late-bound).
- Export with `Extension.SaveAs3` and `swSaveAsOptions_Copy`, so the open document isn't renamed.
- Current method versions: `FeatureExtrusion3`, `FeatureCircularPattern5`,
  `FeatureLinearPattern5`, `InsertMirrorFeature2`, `CreateMate`.
- 60+ exact signatures read from the SOLIDWORKS 2025 type library, each marked [PROVEN] or [DOC].
- Early-binding setup (`gencache` + `cast`) and a late-binding fallback with its quirks.
- Part recipes: verified offset planes, sketch↔model transforms, extrude and cut direction rules,
  loft, sweep, helix, shell, fillet, chamfer, text, equations, fully defined sketches, appearance.
- Assembly workflow: skeleton part, exact placement, mate recipes, `CreateMate` and width mates,
  Toolbox fasteners without size redirection, vendor STEP import, lightweight models, flexible
  sub-assemblies.
- A verification checklist, including interference detection with an "intended overlaps" list.
- Design rules for 3D-printed enclosures.
- Claude Code plugin and marketplace manifests.
- `examples/swlib.py` (Python API layer), two example scripts and a VBA macro.
- `tools/sw_signature.py`, `tools/sw_constants.py`, `tools/sync_skill.py`.
- Documentation and README graphics.
