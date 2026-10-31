# Contributing

Thank you for helping other people automate SOLIDWORKS. Contributions that help most are:

- **A new gotcha with its fix.** Something that failed, why it failed, and the call that works.
- **A tested [DOC] recipe.** You ran it, the checks passed, so it can be marked [PROVEN].
- **Version notes**, e.g. "works on 2023, argument N renamed on 2024".
- **Examples** that build something and check it.

## Rules

1. **No guessed API calls.** Get signatures with `python tools/sw_signature.py IInterface.Method`
   and enum values with `python tools/sw_constants.py swName`, or link the official
   [API help](https://help.solidworks.com) page.
2. **Metres and radians** in every API example. Convert at the boundary (`MM = 0.001`).
3. Mark recipes **[PROVEN]** only after you ran them and checked the result (volume, bounding box,
   What's Wrong = 0, constraint status). Say which SOLIDWORKS version you used.
4. Examples must never touch the user's open documents and never call `CloseAllDocuments`.
   Use `swlib` (early binding) instead of a new connection layer.
5. Keep it public-safe: no personal paths, customer names, part numbers or credentials.
6. Found unknown behaviour? Answer it with a SolidWorks Test Run (`examples/00_test_run_template.py`), not by
   re-running a full build, and add the answer with the SOLIDWORKS version you measured it on.

## Workflow

1. Edit `skills/solidworks-api/SKILL.md`. Never edit the root `SKILL.md`, which is generated.
2. Run:

   ```bash
   python tools/sync_skill.py
   python tools/sync_skill.py --check
   python -m py_compile examples/*.py tools/*.py
   ```

3. Add a line to `CHANGELOG.md`. If the skill changed, bump `version` in both the SKILL.md front
   matter and `.claude-plugin/plugin.json`.
4. Open a pull request that describes what you tested, on which SOLIDWORKS version, and the output.

## Reporting a problem

Open an issue with: your SOLIDWORKS version and service pack, your Python and pywin32 versions,
the failing call, and the full traceback or the What's Wrong messages.
