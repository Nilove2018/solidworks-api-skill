# CLAUDE.md — instructions for Claude working in this repository

This repository is a **Claude Code skill and plugin** that teaches Claude to write correct
SOLIDWORKS API automation (Python + pywin32, and VBA). You are maintaining the skill itself,
its example scripts and its tools. You are not building a CAD model here, unless the user asks.

## Layout

| Path | What it is |
|---|---|
| `skills/solidworks-api/SKILL.md` | **The skill (source of truth).** Claude Code loads it from here |
| `SKILL.md` | Copy of the skill at the root for GitHub readers. Never edit it directly |
| `.claude-plugin/plugin.json`, `marketplace.json` | Plugin and marketplace manifests (the plugin source is the repo root) |
| `examples/swlib.py` | Reusable early-bound API layer used by the examples |
| `examples/01_*.py`, `02_*.py` | Runnable examples (part; skeleton + assembly + checks) |
| `examples/vba/*.bas` | VBA macros to import into a `.swp` |
| `tools/sw_signature.py`, `tools/sw_constants.py` | Read exact signatures and enum values from the installed type library |
| `tools/sync_skill.py` | Copies the skill to the root `SKILL.md` (`--check` to verify) |
| `tools/verify_credit.py` | Checks a signed author credit (INTACT / ALTERED / MISSING); SOLIDWORKS files only with `--sldworks` |
| `docs/` | Installation, usage, troubleshooting |
| `assets/` | SVG graphics for the README |

## Rules when you change the skill

1. Edit `skills/solidworks-api/SKILL.md`, then run `python tools/sync_skill.py`.
2. **Never add a guessed API call or enum value.** Get the signature with
   `python tools/sw_signature.py IInterface.Method` and the value with
   `python tools/sw_constants.py swName`. Cite the API help page if you used one.
3. Mark every recipe **[PROVEN]** (it ran in a real build, with its checks passing) or **[DOC]**
   (taken from the docs or type library and not yet run). Only promote [DOC] to [PROVEN] after a run.
4. Document known limitations and gotchas, with the fix. They save users more time than the API
   docs do.
5. All lengths in API examples are **metres** (`MM = 0.001`), angles in radians. Show
   localised plane names (e.g. Chinese 前视基准面 / 上视基准面 / 右视基准面) only as notes. Code
   finds the planes by type order.
6. Keep it public-safe: no personal paths, company names, part numbers or credentials.
7. Bump `version` in the SKILL.md frontmatter and `.claude-plugin/plugin.json` together, and add a
   `CHANGELOG.md` entry.

## Rules when you change the examples

- Examples must not touch the user's open documents. They create new documents and close only
  files they created themselves (`SwApp.close_path`). Never call `CloseAllDocuments`.
- Each example prints its own checks (volume, bounding box, rebuild errors, constraint status,
  interference) and exits non-zero if a check fails.
- Before saying an example works, run it on a machine with SOLIDWORKS and quote its output.
  `python -m py_compile examples/*.py tools/*.py` is only a syntax check.

## Quick checks before a commit

```bash
python tools/sync_skill.py --check
python -m py_compile examples/*.py tools/*.py
python tools/sw_signature.py IFeatureManager.FeatureExtrusion3
```
