# Examples

All examples create **new** documents and close only files they created themselves. Save your
open work first anyway. Output goes to `examples/out/` (git-ignored).

| File | Builds | Shows |
|---|---|---|
| `swlib.py` | (library) | early-bound COM, model-mm sketching on any plane, verified planes and axes, extrude/cut between coordinates, revolve, helix + sweep, combine, shell, fillet/chamfer by geometry, outline text, equations, configurations, colours, exact extents, export, `Asm` placement + mates (with limits), flexible sub-assemblies, interference, a hidden and signed author credit (`SwApp.credit`, checked with `tools/verify_credit.py`) |
| `00_test_run_template.py` | one answer | a **SolidWorks Test Run**: a 20-line throwaway check in an untitled document (here: the hand of a helix). Copy it to test any unknown behaviour, or the speed of a big sketch, before a full build |
| `01_parametric_plate.py` | 80×50×6 plate, R5 corners, 4 holes, centre boss with bore | sketches on offset planes, through cuts, edge selection by geometry, global variables, volume vs hand calculation, STEP/STL export |
| `02_pinned_assembly.py` | skeleton + base + cover + 2 pins | skeleton datums, exact placement, frame and fastener mate recipes, constraint status, placement deviation, interference detection |
| `vba/CreateFlange.bas` | 25×24×2.5 flange with 4 holes | early-bound VBA with swconst names, finding the Front Plane in any language |

## Run

```bash
pip install pywin32
python examples/01_parametric_plate.py
python examples/02_pinned_assembly.py
```

Each script prints its checks and ends with `ALL CHECKS PASSED` (exit code 0) or
`SOME CHECKS FAILED` (exit code 1).

## Notes

- `swlib.py` is used unchanged from production build scripts, except that the SOLIDWORKS
  version is now detected automatically. It was tested on SOLIDWORKS 2025 SP1.2.
- The two example scripts are new. They are built only from `swlib` calls that were used in
  those production scripts, but the scripts themselves haven't run yet. If one fails on your
  machine, please open an issue with the output.
- The `swlib` 2.1 helpers (configurations, helix, combine, exact extents, `set_global`,
  material fallback, configured components, limit angle mates, mate errors, flexible solving,
  interference) all ran early-bound from this file on SOLIDWORKS 2025 SP1.2, through
  SolidWorks Test Runs on 2026-10-30. The startup wait in `SwApp` hasn't run yet.
- The `swlib` 2.2 helpers (`Asm.mate(verify=)`, save by path, `plane_3pts`, entity names,
  `folder`, `delete_group`, `ray_hits`, `backup`, `log`) passed a SolidWorks Test Run on
  2026-10-31 (14 of 14 checks).
- The author-credit helpers (`stamp_credit`, `read_credit`, `stamp_file_ads`, `stamp_step`,
  `stamp_pdf`) stamped a part, an assembly, a drawing, a STEP and a PDF, and all five checked
  INTACT (2026-10-31). Check SOLIDWORKS files from a second process only with
  your work saved: one such check was followed by a SOLIDWORKS crash (SKILL.md §8).
- Set `SW_TLB_MAJOR=32` (2024) or `31` (2023) to force an older type library.

## VBA

A `.swp` file is binary, so it is shipped as `.bas` source:
1. SOLIDWORKS > **Tools > Macro > New…**, save as `CreateFlange.swp`
2. In the VBA editor: **File > Import File…** > `CreateFlange.bas`
3. **Tools > References**: tick the SOLIDWORKS Type Library and the SOLIDWORKS Constant type library
4. Run with F5, or **Tools > Macro > Run**
