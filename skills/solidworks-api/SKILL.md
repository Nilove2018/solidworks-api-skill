---
name: solidworks-api
description: >
  SOLIDWORKS API automation skill (SOLIDWORKS 2025, Python pywin32 COM, VBA macros). Builds
  parametric parts, fully defined sketches, skeleton-based assemblies with mates, fasteners,
  vendor STEP imports, interference checks, STEP/STL/PNG exports. Encodes exact API signatures,
  correct swconst values and lessons from real production builds (configurations, limit angle
  mates, flexible sub-assemblies, helix springs, multi-body combine, toleranced drawings with GD&T,
  safe edits of existing parts, ray-cast measuring).
  Triggers: SolidWorks, SOLIDWORKS, SLDWORKS, SldWorks.Application, ISldWorks, IModelDoc2,
  IFeatureManager, ISketchManager, IAssemblyDoc, AddMate5, CreateMate, sketch, extrude, revolve,
  loft, sweep, helix, spring, torsion spring, cut, fillet, chamfer, shell, pattern, mirror, combine,
  configuration, mate, limit mate, flexible, assembly, .SLDPRT, .SLDASM, .swp, VBA macro, swconst,
  Toolbox, enclosure, 3D-print part, drawing, SLDDRW, GD&T, tolerance, dimension, vent, existing part,
  test run, cutterhead, 刀盘.
version: 2.3.0
author: Nilove Mandal
tags: [solidworks, cad, com-api, pywin32, vba, automation, mechanical-design, assembly, 3d-printing]
---

# SOLIDWORKS API Skill (our version)

Version 2 of the community "SolidWorks API skill" (originally written in Chinese for SOLIDWORKS
2023 and a shield-machine cutterhead project). This version is translated into English, has its
constant values corrected against the SOLIDWORKS 2025 type library, and adds what we learned
building real parts and assemblies through the API in 2026: 3D-printed enclosures, a
59-component product assembly with 174 mates, optics shrouds, springs and flexible sub-assemblies.

Status markers used below:
- **[PROVEN]**: used in one of our working build scripts on SOLIDWORKS 2025 SP1.2. The
  torsion-spring items ran late-bound in that build. Their early-bound forms in `swlib.py` 2.1
  ran on 2026-10-30. Where the two modes are written differently, section 3.0 gives both.
- **[DOC]**: taken from the official API help or the installed type library and not yet run by us.
  Check the result after the call (see section 12) before you rely on it.

## 0. Read this first: what cost the most failed runs

Three builds show where the time goes. Each row is now a rule in this file.
- **Torsion spring** (a part with configurations, a flexible sub-assembly and a test rig): 28 of
  about 50 script runs failed or gave a wrong result.
- **Vernier caliper** (parts, assembly, 7 toleranced drawings): 3.5 h, of which about 2 h could
  have been avoided.
- **Vent grilles added to an existing housing** (28 light-tight slots): about 45 min went on 9
  full rebuilds that tested small logic changes.

| Cause | Cost | Rule |
|---|---|---|
| Speed never tested: 201 engraved lines plus digits in **one sketch**, all fixed (20 min to solve); a drawing script that walked every face and edge of a part with 5,000 faces (13 min, plus 2 stopped runs) | ~1 h | 1.14, 7.12, 8.2 |
| The full build re-run to test one small change (28 slots, 4–7 min each) | ~45 min | 1.14, 11 |
| Drawing layout fixed by trial: views off the sheet, oversized text, BOM over the title block, texts on texts | 6 reruns, ~20 min | 8.3 |
| Design choice not confirmed (recessed instead of flush); pattern not laid out around the feature that interrupts it | 2 redesigns | 1.16, 14.3 |
| Fit and clearance table done *after* the parts and drawings existed | rod, assembly and every drawing rebuilt | 1.17 |
| Existing part: geometry scanned before the user's last edit, hidden internal rail found only while cutting, added material too close to a drafted skin | 3 dry runs | 1.18, 9.1 |
| A mate solved to −40 mm instead of +40 mm; a save failed because the old file was still open under its old title | 2 reruns | 10.3, 8 |
| A new late-bound helper layer was written instead of importing `swlib.py`. Calls then behaved differently from one object to the next | 11 failed runs | 1.13, 3.0 |
| A new feature only *touched* the body (face to face, or tangent), so the merge failed | 3–4 failed runs | 7.6 |
| Configurations: suppression didn't apply, parents took their children with them, and a temporary suppress left the part in 2 bodies | 3 failed runs | 7.11 |
| A limit angle mate whose range included 0°, a drag test using `Transform2`, `FeatureByPositionReverse(0)` used in an assembly | 4 failed runs | 10.3, 10.8, 1.6 |
| The whole build was re-run to answer one question about unknown behaviour | most of the time | 1.14 |
| Build scripts patched with shell heredocs: strings broken, a file truncated to 0 bytes, a replace silently skipped | 6 failed runs | 1.15 |
| An ambiguous drawing letter was misread, and everything was rebuilt | 1 full rebuild | 1.16, 15 |
| Progress hidden (`\| tail`), blocked `sleep`, a kill that missed the real Python process | ~10 min | 11 |

---

## 1. Ground rules (read these first)

1. **Never guess an API call.** SOLIDWORKS methods take 6–27 positional arguments of mixed types.
   Use the exact signatures in section 6. For anything not listed, read the signature from the
   generated type library (section 3.3) or help.solidworks.com before writing the call. Pass
   every argument explicitly.
2. **Never hard-code enum integers.** Load them from `swconst` (section 3.2). The original skill
   had wrong values (section 5.2). Literal numbers are OK only when commented with the enum name.
3. **All API lengths are metres, angles are radians.** Keep the design in mm and convert at the
   API boundary with `MM = 0.001`. Set the document to MMGS so the user sees mm.
4. **Respect the user's SOLIDWORKS session.**
   - Never call `CloseAllDocuments`. It closes every document without saving. (This once wiped a
     user's open work.) Close only documents your script created or that live under the project
     folder.
   - Don't re-run builds in the live session while the user is working in it. Ask first, and say
     "please don't click in SOLIDWORKS for the next minute" before a run.
   - Save and restore every user preference you change (try/finally or `atexit`).
5. **Select by geometry or by name, never by screen picks.** Find faces and edges by their normal
   and position (section 9). Find features with `FeatureByName` on names you gave them.
   `FeatureByName` lives on `IPartDoc`, `IAssemblyDoc` and `IComponent2`, **not on `IModelDoc2`**,
   so cast the document first.
6. **Name everything.** Every sketch, feature, plane, axis and mate gets a readable name right
   after it is created. In a **part**, the new feature is `FeatureByPositionReverse(0)`. In an
   **assembly** it is not, because the Mates folder sits after the reference geometry, so
   `FeatureByPositionReverse(0)` returns the MateGroup. There, walk the tree and take the last
   feature of the expected type (`newest(doc, "RefAxis")`, section 10.3). Mates are the last
   sub-feature of the MateGroup.
7. **Check every step.** A `None` return means failure, so raise immediately with the feature name.
   A non-`None` feature can still be wrong (it went in the wrong direction or cut nothing), so
   check volume, bounding box and body count after each feature (section 12).
8. **One script rebuilds a part from zero.** All dimensions live in one parameters file. To change
   the design, edit the parameters and re-run. Don't hand-patch saved files.
9. **Parametric output.** Put every size into Global Variables (Tools > Equations). Fully define
   every sketch, either with dimensions linked to the variables (preferred) or with Fixed
   relations. The user must be able to edit a variable and press Ctrl+B.
10. **Assemblies: skeleton plus mates.** The first component is a skeleton part of named datum
    planes and axes. Mate every other component to it so it is fully constrained. Check with
    What's Wrong = 0, all components "fully constrained" and Interference Detection.
11. **Check the input dimensions before you model.** If the given numbers can't be built (for
    example a flange smaller than the wall on it, or holes breaking into an opening), say so,
    propose the fix, and list every change in the report.
12. **Report plainly.** Report what was built, what was checked (with numbers), what was assumed,
    and what the user must measure or verify.
13. **Import `swlib.py`. Don't write a new connection layer.** Copy it into the project's
    `Scripts/` folder and add what you need to it. Use one binding mode per script, early binding,
    through its helpers (section 3.0). A fresh hand-written late-bound layer cost 11 failed runs
    in one build. Old project scripts are not a template: check which binding they use before you
    copy anything from them.
14. **Test run first, then build (mandatory).** A **SolidWorks Test Run** (short: *test run*) is
    a 20–40 line throwaway script that checks one API behaviour on a dummy part (for example a
    60 × 20 × 5 mm block) in a new, untitled document. It prints the answer, closes only that
    document, and the answer goes into this skill before the build script is written or changed.
    No build script may use a behaviour that isn't marked [PROVEN] or confirmed by a test run in
    this session: which side a plane lands on, a helix's hand, a mate's direction, whether
    touching features merge, configuration suppression, every drawing call. **This includes
    speed**: before any sketch with more than 40 segments, or any loop over the faces and edges of
    a detailed part, time a small test run and extrapolate. Never re-run a 1–7 minute build to
    answer one question; test new feature logic on **one** instance (one slot, one grille) on a
    scratch copy first. The template is `examples/00_test_run_template.py`.
15. **Edit scripts with the Edit tool, never with shell string replacement.** Patching a build
    script through `python - <<'EOF'` find-and-replace gave `SyntaxError: unterminated string
    literal` twice, because a `"\n"` in the patch became a real newline. Heredoc patches with
    non-ASCII text (≈ ² ·) also truncated a file to 0 bytes and silently skipped a replace. If a
    script must patch text: `encoding="utf-8"`, `assert old in s` before every replace, and build
    the new text completely before opening the file for writing. In code that generates code,
    write `chr(10)` instead of `"\n"`. Before any run that changes a real part:
    `python -m py_compile` and a dry mode with no SOLIDWORKS calls, so a missing helper can't crash
    the script *after* a destructive step. Never name a script like a standard-library module
    (`inspect.py` broke numpy and scipy). Write Windows paths as raw strings
    (`r"D:\Parts\test.SLDPRT"`) or with forward slashes: in a normal string `\t`, `\b`, `\n`, `\a`,
    `\f`, `\v` become control characters, so the path names another file or generated text gets an
    invisible tab or backspace (it happened in two generated paths). Before writing generated text,
    `re.search(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", s)` must be None.
16. **Confirm the interpretation before you build everything.** When a drawing letter, a note or a
    feature can be read two ways, build only the part first and send the user a render
    (`SendUserFile`). Get a yes before you build sub-assemblies and test rigs on top of it. The same
    applies to style: recessed, flush or chamfered, and the layout of a pattern. Show a quick
    render and get a yes before any CAD.
17. **Fits, clearances and standard values before modelling.** Write the parameters file first:
    every sliding or press fit gets a min/max clearance check, and every value taken from a
    standard (accuracy class, tolerance grade, thread data) is **computed there**, never typed
    into a note by hand. A fit changed after the drawings existed (3.90 → 3.990) forced a rebuild
    of the part, the assembly and every drawing. A hand-typed accuracy note was wrong, because
    ISO 13385-1 rounds *up* to 10 µm.
18. **Editing a part that already exists** (the user's file):
    1. Back it up first (`_Backup\<part>_BACKUP_<date>_<reason>.SLDPRT`, verify the MD5), and again
       before every later revision.
    2. Note the file's modified time. Right before cutting, check it again and re-measure the part
       if it changed: the user may have edited it since your last scan.
    3. Check the session: `GetOpenDocumentByName`, `GetSaveFlag()` (unsaved user edits?), and scan
       the feature names to prove no user edits are mixed in with yours.
    4. Measure the real part before designing (9.1): the outer skin, wall thickness, and the material
       *behind* every planned footprint.
    5. Dry run on a **renamed copy** (`DRYRUN<n>_<part>.SLDPRT`). SOLIDWORKS can't open two files
       with the same name.
    6. Name every new feature `<GROUP>-<ID> ...` and put each group in a tree folder
       (`InsertFeatureTreeFolder2`). A revision is then "delete everything starting with `VENT-V2`"
       (`DeleteSelection2(swDelete_Absorbed | swDelete_Children)`), and the volume returns exactly to
       the original. Build and remove each group with its own function: rebuilding one group took
       ~1 min, the whole part 4–7 min.

---

## 2. Environment (check these on your machine)

| Item | Typical value |
|---|---|
| SOLIDWORKS | Written and tested on 2025 SP1.2 (type library major **33**). 2024 = 32, 2023 = 31. Most calls work back to ~2018 |
| Install dir | `C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS` (`sldworks.tlb`, `swconst.tlb`, `swcommands.tlb`) |
| ProgID | `SldWorks.Application` (version-specific: `SldWorks.Application.33`) |
| Toolbox data | `C:\SOLIDWORKS Data\browser\<standard>\...`. Ask the user or read Toolbox settings, because the location varies |
| Python | 3.9+ with `pywin32` (`pip install pywin32`). If pywin32 is in the *user* site, don't run scripts with `python -I`, because it hides the user site |
| makepy cache | `%LOCALAPPDATA%\Temp\gen_py\<py-ver>\83A33D31-27C5-11CE-BFD4-00400513BB57x0x33x0.py` (sldworks) and `4687F359-55D0-4CD3-B6CF-2EB42C11F989x0x33x0.py` (swconst), created by `gencache.EnsureModule` |
| Add-ins | SOLIDWORKS CAM and other add-ins can pop modal dialogs that block COM calls (section 11) |

Templates: never hard-code a template path. Use
`app.GetUserPreferenceStringValue(swDefaultTemplatePart / Assembly / Drawing)` (8 / 9 / 10).
The original skill's path `D:\SW2023\...\chinese-simplified\tutorial\part.prtdot` was specific to
one machine.

---

## 3. Connecting from Python

### 3.0 Binding rules (read before you write any COM call) [PROVEN]

pywin32 talks to SOLIDWORKS in two modes. **Early-bound** objects are makepy wrappers
(`type(obj).__module__` contains `gen_py`). **Late-bound** objects are `CDispatch`. The same
member behaves differently in the two modes. Mixing them is the single biggest source of failed
runs. You get a mix when you attach late-bound (`GetActiveObject`, `Dispatch`) while a makepy
cache exists: some returned objects (the result of `OpenDoc6`, and anything taken from it) come
back early-bound anyway.

| Member | Early-bound | Late-bound |
|---|---|---|
| Zero-argument method (`GetTitle`, `GetSketchSegments`, `GetConfigurationNames`, `GetChildren`, `GetConstrainedStatus`, `GetCount`, `CreateMassProperty`) | `obj.GetTitle()` | evaluated as a property: `obj.GetTitle` (with `()` it fails: *'tuple' / 'int' object is not callable*). Always use `call(obj, "GetTitle")` |
| ByRef output (`OpenDoc6` Errors/Warnings, `GetErrorCode2`, `GetExtremePoint`, `GetMaterialPropertyName2`) | leave out pure outputs, or pass plain `0` for in/out ones; **read the returned tuple**: `ok, x, y, z = body.GetExtremePoint(dx, dy, dz)` | pass `VARIANT(VT_BYREF \| VT_I4, 0)` (or `VT_R8`, `VT_BSTR`, `VT_BOOL`) and read `.value`. Passing a VARIANT early-bound fails: *int() argument must be ... not 'VARIANT'* |
| "Nothing" for an argument typed as an object (`ICallout`, `IBody2 MainBody`) | `None` | `NOTHING = VARIANT(VT_DISPATCH, None)`. Python `None` gives *Type mismatch* |
| "Nothing", or an array of objects, for an argument typed `variant` (`ToolVar`, `EntitiesToMate`) | `NOTHING` / `dispatch_array(objs)` | the same |
| `FeatureByName` | on `IPartDoc` / `IAssemblyDoc` / `IComponent2` only; `IModelDoc2` has none (*Member not found*) | resolves on any document |
| Property put that makepy marks read-only (`IComponent2.Solving`) or that takes an index (`IEquationMgr.Equation(i)`) | `com_put(obj, "Solving", 1)` | `com_put(...)`. Late-bound pywin32 has no syntax for an indexed put |
| `IMathUtility.CreatePoint` / `CreateTransform` | works | throws *The server threw an exception*. Use the by-hand transform (7.3) or cast `IMathUtility` early-bound |

```python
NOTHING = VARIANT(pythoncom.VT_DISPATCH, None)

def dispatch_array(objs):           # array of COM objects for a `variant` argument
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, list(objs))

def com_put(obj, name, *args):      # property put through IDispatch, works in both modes
    o = obj._oleobj_
    o.Invoke(o.GetIDsOfNames(name), 0, pythoncom.DISPATCH_PROPERTYPUT, 0, *args)

def is_early(obj):
    return "gen_py" in type(obj).__module__
```

Rules:
1. Use early binding through `swlib.py` (`cast()` on every returned object). Don't attach
   late-bound "to be quick".
2. If you are forced to use late binding, then *every* zero-argument member goes through
   `call(obj, "Name")`, every Nothing is `NOTHING`, and every ByRef is a `VARIANT`. Never write
   `obj.Name()` directly.
3. Read the argument types before the call: `python tools/sw_signature.py IPartDoc.GetMaterialPropertyName2`
   shows `ByRef string Database`, which tells you it is an output.

### 3.1 Early binding (preferred) [PROVEN]

Early binding makes SOLIDWORKS methods behave like methods. Late binding confuses zero-argument
methods with properties.

```python
import pythoncom, win32com.client as w32
from win32com.client import VARIANT, gencache

SW  = gencache.EnsureModule('{83A33D31-27C5-11CE-BFD4-00400513BB57}', 0, 33, 0)            # sldworks
SWC = gencache.EnsureModule('{4687F359-55D0-4CD3-B6CF-2EB42C11F989}', 0, 33, 0).constants  # swconst
MM = 0.001

def cast(obj, iface):
    """Return obj as the early-bound wrapper of interface `iface` (QueryInterface)."""
    if obj is None:
        return None
    cls = getattr(SW, iface)
    o = obj._oleobj_ if hasattr(obj, "_oleobj_") else obj
    try:
        o = o.QueryInterface(cls.CLSID, pythoncom.IID_IDispatch)
    except pythoncom.com_error:
        pass
    return cls(o)

def darr(vals):   # Python list -> SAFEARRAY(double) for array arguments
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(v) for v in vals])

try:
    raw = w32.GetActiveObject("SldWorks.Application")   # attach to the running session
except pythoncom.com_error:
    raw = w32.Dispatch("SldWorks.Application")          # or start it (takes ~1 min)
app = cast(raw, "ISldWorks"); app.Visible = True
```

**Starting SOLIDWORKS from a script**: while the splash screen shows, every call fails with
*The interface is unknown* (-2147023179). We saw this on 2026-10-30. Retry `GetActiveObject` plus
one cheap call every few seconds until it answers (up to ~4 min). `swlib.SwApp()` does this
[DOC: the retry loop itself hasn't run yet].

**Inside CAD Master** (the add-in that runs Claude Code in a SOLIDWORKS task pane) [PROVEN 2026-10-31]:
`CADMASTER_SW_PID` names the SOLIDWORKS that hosts the panel. `swlib.SwApp()` (2.2+) attaches to it
through the ROT moniker `SolidWorks_PID_<pid>` (`swlib.cadmaster_instance()`); `GetActiveObject` may
return another instance (a hidden `-Embedding` SOLIDWORKS was running on 2026-10-31). Without swlib:
`import cadmaster_sw` (on PYTHONPATH in CAD Master) and `cast(cadmaster_sw.get_raw(), "ISldWorks")`.

Every object an API call returns must be cast before use: `cast(doc, "IModelDoc2")`,
`cast(doc.FeatureManager, "IFeatureManager")`, `cast(f, "IFeature")`, `cast(c, "IComponent2")`
and so on. ByRef output arguments come back as a tuple: `doc, err, warn = app.OpenDoc6(...)`.

### 3.2 Late binding (fallback) [PROVEN]

Use late binding only if early binding fails. It failed once, in a session where SOLIDWORKS was
started by `Dispatch`. If you do, follow every rule in 3.0, because objects can still come back
early-bound.

```python
sw = win32com.client.Dispatch("SldWorks.Application")

def call(obj, name):
    """Late-bound COM exposes zero-arg SOLIDWORKS methods as properties."""
    v = getattr(obj, name)
    return v() if callable(v) and not isinstance(v, win32com.client.CDispatch) else v

def sw_constants(sw_dir=r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS"):
    """All swconst enum values straight from the type library - never guessed."""
    tlb = pythoncom.LoadTypeLib(os.path.join(sw_dir, "swconst.tlb"))
    out = {}
    for i in range(tlb.GetTypeInfoCount()):
        ti = tlb.GetTypeInfo(i); ta = ti.GetTypeAttr()
        if ta.typekind == pythoncom.TKIND_ENUM:
            for j in range(ta.cVars):
                vd = ti.GetVarDesc(j); out[ti.GetNames(vd.memid)[0]] = vd.value
    return out
C = sw_constants()
```

Late-binding quirks:
- Zero-argument methods (`GetTitle`, `FirstFeature`, `GetNextFeature`, `GetTypeName2`,
  `EditRebuild3`, `ViewZoomtofit2`, `GetFaces`, `GetBox`, `GetEquationMgr`, `CreateMassProperty`,
  `GetSketchSegments`, `GetConfigurationNames`, `GetChildren`, `GetConstrainedStatus`,
  `EditSuppress2`, `GetCount`) must go through `call()`.
- Every `Nothing` is `NOTHING` (Callout in `SelectByID2`, Data in `Select4`, MainBody in
  `InsertCombineFeature`). Every array of objects is `dispatch_array(...)`.
- A ByRef int output (the AddMate5 error) is `VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)`;
  read it back with `.value`. A ByRef string (Database in `GetMaterialPropertyName2`) is
  `VARIANT(VT_BYREF | VT_BSTR, "")`, and a ByRef bool (IsWarning in `GetErrorCode2`) is
  `VARIANT(VT_BYREF | VT_BOOL, False)`.
- `body.GetExtremePoint(x, y, z)` returns `(ok, x, y, z)`.
- `IMathUtility.CreatePoint(...).MultiplyTransform(...)` throws. Use the transform by hand (7.3).
- Correction to the original skill: "methods with more than 12 parameters fail via IDispatch"
  is **not** true for pywin32. `FeatureExtrusion2` (23 args) and `FeatureCut4` (27 args) worked
  late-bound from Python. That limitation applies to Node.js `winax`, where the `SelectByID2`
  Callout argument is also rejected (use `FeatureByName()` + `Select2()` there).

### 3.3 Reading an exact signature yourself

Search the makepy file for `\tdef <Method>(self` inside the right `class I<Interface>`. The
argument names are the official ones. To list enum values, grep the swconst makepy file for
`^\s+swSomething\s*=`. Online reference: `help.solidworks.com/2025/english/api/sldworksapi/...`.
The pages render with JavaScript, so read them in a browser, not with a plain HTTP fetch.

---

## 4. Object model (short)

```
ISldWorks (application)
├── IModelDoc2 ── IPartDoc (.sldprt) | IAssemblyDoc (.sldasm) | IDrawingDoc (.slddrw)
│   │                (FeatureByName is on IPartDoc / IAssemblyDoc, GetBodies2 on IPartDoc; neither on IModelDoc2)
│   ├── .Extension        IModelDocExtension   SelectByID2, SaveAs3, CreateMassProperty, custom props
│   ├── .SketchManager    ISketchManager       InsertSketch, Create* geometry, AddToDB
│   ├── .FeatureManager   IFeatureManager      Feature*, Insert*, CreateDefinition/CreateFeature
│   ├── .SelectionManager ISelectionMgr        CreateSelectData, GetSelectedObject6
│   ├── GetEquationMgr()  IEquationMgr         global variables / equations
│   └── ActiveView        IModelView           EnableGraphicsUpdate
├── IFeature  (GetTypeName2, Name, GetFaces, GetSpecificFeature2, GetDefinition/ModifyDefinition)
├── IBody2 → IFace2 → IEdge / ISurface / ICurve            (B-rep topology)
├── IComponent2 (assembly instance: Name2, Transform2, FeatureByName, GetConstrainedStatus,
│                ReferencedConfiguration, Solving = rigid/flexible)
└── IMathUtility (CreatePoint, CreateTransform, CreateVector)
```

---

## 5. Constants

### 5.1 Correct values (SOLIDWORKS 2025 swconst, read from the type library)

| Enum | Values |
|---|---|
| swDocumentTypes_e | swDocPART=1, swDocASSEMBLY=2, swDocDRAWING=3 |
| swEndConditions_e | **swEndCondBlind=0**, **swEndCondThroughAll=1**, swEndCondThroughNext=2, swEndCondUpToVertex=3, swEndCondUpToSurface=4, swEndCondOffsetFromSurface=5, swEndCondMidPlane=6, swEndCondUpToBody=7, swEndCondThroughAllBoth=9 |
| swUserPreferenceStringValue_e | swDefaultTemplatePart=8, swDefaultTemplateAssembly=9, swDefaultTemplateDrawing=10 |
| swUserPreferenceToggle_e | swInputDimValOnCreate=10 (dimension pop-up), swSketchInference=249 (snapping), swViewDisplayHideAllTypes=198, swMultiCAD_Enable3DInterconnect=691 |
| swUserPreferenceIntegerValue_e | swUnitSystem=263 (value swUnitSystem_MMGS=5), swStepAP=75 |
| swRefPlaneReferenceConstraints_e | Parallel=1, Perpendicular=2, Coincident=4, **Distance=8**, Angle=16, MidPlane=128, **OptionFlip=256** |
| swMateType_e | COINCIDENT=0, CONCENTRIC=1, PERPENDICULAR=2, PARALLEL=3, TANGENT=4, DISTANCE=5, ANGLE=6, SYMMETRIC=8, GEAR=10, WIDTH=11, LOCK=16, SCREW=17, SLOT=21, HINGE=22 |
| swMateAlign_e | ALIGNED=0, ANTI_ALIGNED=1, CLOSEST=2 |
| swAddMateError_e | ErrorUknown=0, **NoError=1**, IncorrectMateType=2, IncorrectAlignment=3, IncorrectSelections=4, OverDefinedAssembly=5, IncorrectGearRatios=6 |
| swConstrainedStatus_e | UnderConstrained=2, **FullyConstrained=3**, OverConstrained=4 (1 = unknown/fixed) |
| swSaveAsOptions_e | Silent=1, Copy=2, SaveReferenced=4, UpdateInactiveViews=16. swSaveAsCurrentVersion=0 |
| swOpenDocOptions_e | Silent=1, ReadOnly=2, LoadModel=16 |
| swSelectType_e | EDGES=1, FACES=2, VERTICES=3, DATUMPLANES=4, DATUMAXES=5, DATUMPOINTS=6, SKETCHES=9, SKETCHSEGS=10, SKETCHPOINTS=11, COMPONENTS=20, MATES=21, BODYFEATURES=22, EXTSKETCHPOINTS=25, HELIX=26, COORDSYS=61, SOLIDBODIES=76 |
| swBodyOperationType_e | SWBODYINTERSECT=15901, SWBODYCUT=15902, **SWBODYADD=15903** (`InsertCombineFeature`) |
| swInConfigurationOpts_e | swThisConfiguration=1, swAllConfiguration=2, swSpecifyConfiguration=3 |
| swFeatureSuppressionAction_e | swSuppressFeature=0, swUnSuppressFeature=1 |
| swComponentSolvingOption_e | swComponentRigidSolving=0, **swComponentFlexibleSolving=1** |
| swAddComponentConfigOptions_e | CurrentSelectedConfig=0 (use the config named in ExistingConfigName), NewConfigWithAllReferenceModels=1, NewConfigWithAsmStructure=2 |
| misc | swSolidBody=0, swFmSweep=17, swFmHoleWzd=25, swFmLPattern=6, swFmCirPattern=5, swHelixDefinedByHeightAndRevolution=1, swChamferAngleDistance=1, swFeatureTreeFolder_Containing=2, swNotAToolboxPart=0, swSetValue_InAllConfigurations=2, swCustomInfoText=30, swCustomPropertyReplaceValue=2, swMateWidth_Centered=0, swRunMacroDefault=0, swRunMacroUnloadAfterRun=1 |

### 5.2 Corrections to the original skill

| Original said | Correct | Effect of the old value |
|---|---|---|
| swEndCondBlind = 1 | **0** | 1 is ThroughAll, so a "blind" extrude went through everything |
| swEndCondThroughAll = 2 | **1** | 2 is ThroughNext |
| swSelEDGES = 4 | **1** | 4 is DATUMPLANES |
| swSelSKETCHES = 16 | **9** | |
| swPlaneTop/Front/Right = 1/2/3 | no such enum | Find the planes by type order (section 7.1) |
| `FeatureExtrusion2` "fails >12 params" | works in pywin32; prefer `FeatureExtrusion3` | see 3.2 |
| `part.SaveAs3(path + ".STEP", 0, 0)` | `Extension.SaveAs3(path, 0, swSaveAsOptions_Silent \| swSaveAsOptions_Copy, None, None, 0, 0)` | without Copy the document is renamed to the export |
| `FeatureCircularPattern4`, `FeatureLinearPattern4`, `FeatureMirror2` | `FeatureCircularPattern5`, `FeatureLinearPattern5`, `InsertMirrorFeature2` | newer signatures, see 6 |

### 5.3 Corrections to version 2.0 of this skill (found in the torsion-spring build)

| v2.0 said | Correct | Section |
|---|---|---|
| Name a new feature through `FeatureByPositionReverse(0)` | Only in a part. In an assembly that returns the MateGroup; use `newest(doc, type)` | 1.6, 10.3 |
| 7.3 sketch transform through `IMathUtility.CreatePoint` | That throws when late-bound. The by-hand transform works in both modes | 7.3 |
| `mate()` helper passed 0 for all three angle arguments | Positions 9–11 of AddMate5 are `Angle, AngleAbsUpperLimit, AngleAbsLowerLimit`, in radians | 10.3 |
| Flexible sub-assembly: "tell the user to right-click" | `com_put(comp, "Solving", swComponentFlexibleSolving)`, then read it back | 10.7 |
| `AddComponent5(path, 0, "", False, "", ...)` only | Pass the configuration name as ExistingConfigName, then check `ReferencedConfiguration` | 10.2 |
| Appearance set on "the whole part" | Appearance is per configuration: set it in each one | 7.10 |
| Size checks with `GetBodyBox` / `GetPartBox` | Those are loose on curved bodies. Use `GetExtremePoint` in 6 directions | 9, 12 |

---

## 6. Exact signatures (SOLIDWORKS 2025, from the installed type library)

Arguments in order. `[n]` = number of arguments. Length arguments are in metres.

**ISldWorks**
- `NewDocument(TemplateName, PaperSize, Width, Height)` [4]
- `OpenDoc6(FileName, Type, Options, Configuration, Errors, Warnings)` [6]
- `LoadFile4(FileName, ArgString, ImportData, Errors)` [4], with `GetImportFileData(FileName)` for STEP/IGES import
- `ActivateDoc3(Name, UseUserPreferences, Option, Errors)` [4] · `CloseDoc(Name)` [1]
- `RunMacro2(FilePathName, ModuleName, ProcedureName, Options, Error)` [5]

**IModelDocExtension**
- `SelectByID2(Name, Type, X, Y, Z, Append, Mark, Callout, SelectOption)` [9]
- `SaveAs3(Name, Version, Options, ExportData, AdvancedSaveAsOptions, Errors, Warnings)` [7] (returns bool)
- `SetUserPreferenceInteger(UserPref, UserPrefOption, Value)` [3] · `DeleteSelection2(DeleteOptions)` [1]
- `CreateMassProperty()` · `GetWhatsWrongCount()` · `CustomPropertyManager(Config)`

**ISketchManager**
- `InsertSketch(UpdateEditRebuild)` [1]: opens a sketch on the selected plane or face, or closes the open sketch
- `CreateLine(X1, Y1, Z1, X2, Y2, Z2)` [6] · `CreateCenterLine(...)` [6]
- `CreateCornerRectangle(X1, Y1, Z1, X2, Y2, Z2)` [6] (returns 4 segments)
- `CreateCenterRectangle(Xc, Yc, Zc, Xcorner, Ycorner, Zcorner)` [6] (returns the edges **plus construction diagonals**)
- `CreateCircleByRadius(XC, YC, Zc, Radius)` [4]
- `CreateArc(XC, YC, Zc, X1, Y1, Z1, X2, Y2, Z2, Direction)` [10], Direction +1 = CCW, −1 = CW in *sketch* space
- `CreatePolygon(XC, YC, Zc, Xp, Yp, Zp, Sides, Inscribed)` [8] [DOC]
- `CreateSketchSlot(SlotCreationType, SlotLengthType, Width, X1..Z3, CenterArcDirection, AddDimension)` [14] [DOC]
- `CreateSpline2(PointData, SimulateNaturalEnds)` [2] [DOC] · `CreatePoint(X, Y, Z)` [3]
- `SketchOffset(Offset, BothDirections, Chain, CapEnds, MakeConstruction, AddDimensions)` [6] [DOC]
- `SketchTrim(Option, X, Y, Z)` [4] [DOC]

**IModelDoc2**
- `AddDimension2(X, Y, Z)` [3] (also `AddHorizontalDimension2` / `AddVerticalDimension2`): the text position; dimensions the selection
- `SketchAddConstraints(IdStr)` [1]: relation on the selection, e.g. `"sgFIXED"`, `"sgCOINCIDENT"`, `"sgATMIDDLE"`
- `InsertSketchText(Ptx, Pty, Ptz, Text, Alignment, FlipDirection, HorizontalMirror, WidthFactor, SpaceBetweenChars)` [9]
- `InsertHelix(Reversed, Clockwised, Tapered, Outward, Helixdef, Height, Pitch, Revolution, TaperAngle, Startangle)` [10]
- `InsertAxis2(AutoSize)` [1] (two planes selected) · `InsertFeatureShell(Thickness, Outward)` [2]
- `FeatureByPositionReverse(Num)` [1] · `ForceRebuild3(TopOnly)` [1] · `ShowNamedView2(VName, ViewId)` [2] · `SaveBMP(File, W, H)` [3] · `Parameter(Name)` [1]

**IFeatureManager**
- `FeatureExtrusion3(Sd, Flip, Dir, T1, T2, D1, D2, Dchk1, Dchk2, Ddir1, Ddir2, Dang1, Dang2, OffsetReverse1, OffsetReverse2, TranslateSurface1, TranslateSurface2, Merge, UseFeatScope, UseAutoSelect, T0, StartOffset, FlipStartOffset)` [23]
- `FeatureCut4(Sd, Flip, Dir, T1, T2, D1, D2, Dchk1, Dchk2, Ddir1, Ddir2, Dang1, Dang2, OffsetReverse1, OffsetReverse2, TranslateSurface1, TranslateSurface2, NormalCut, UseFeatScope, UseAutoSelect, AssemblyFeatureScope, AutoSelectComponents, PropagateFeatureToParts, T0, StartOffset, FlipStartOffset, OptimizeGeometry)` [27]
- `FeatureRevolve2(SingleDir, IsSolid, IsThin, IsCut, ReverseDir, BothDirectionUpToSameEntity, Dir1Type, Dir2Type, Dir1Angle, Dir2Angle, OffsetReverse1, OffsetReverse2, OffsetDistance1, OffsetDistance2, ThinType, ThinThickness1, ThinThickness2, Merge, UseFeatScope, UseAutoSelect)` [20]
- `InsertProtrusionBlend2(Closed, KeepTangency, ForceNonRational, TessToleranceFactor, StartMatchingType, EndMatchingType, StartTangentLength, EndTangentLength, StartTangentDir, EndTangentDir, IsThinBody, Thickness1, Thickness2, ThinType, Merge, UseFeatScope, UseAutoSelect, GuideCurveInfluence)` [18] (loft)
- `InsertCutBlend(Closed, KeepTangency, ForceNonRational, TessToleranceFactor, StartMatchingType, EndMatchingType, IsThinBody, Thickness1, Thickness2, ThinType, UseFeatScope, UseAutoSelect)` [12] (loft cut)
- `FeatureFillet3(Options, R1, R2, Rho, Ftyp, OverflowType, ConicRhoType, Radii, Dist2Arr, RhoArr, SetBackDistances, PointRadiusArray, PointDist2Array, PointRhoArray)` [14]
- `InsertFeatureChamfer(Options, ChamferType, Width, Angle, OtherDist, VertexChamDist1, VertexChamDist2, VertexChamDist3)` [8]
- `InsertRefPlane(FirstConstraint, FirstValue, SecondConstraint, SecondValue, ThirdConstraint, ThirdValue)` [6]
- `FeatureCircularPattern5(Number, Spacing, FlipDirection, DName, GeometryPattern, EqualSpacing, VaryInstance, SyncSubAssemblies, BDir2, BSymmetric, Number2, Spacing2, DName2, EqualSpacing2)` [14] [DOC]
- `FeatureLinearPattern5(Num1, Spacing1, Num2, Spacing2, FlipDir1, FlipDir2, DName1, DName2, GeometryPattern, VaryInstance, HasOffset1, HasOffset2, CtrlByNum1, CtrlByNum2, FromCentroid1, FromCentroid2, RevOffset1, RevOffset2, Offset1, Offset2, D2PatternSeedOnly, SyncSubAssemblies)` [22] [DOC]
- `InsertMirrorFeature2(BMirrorBody, BGeometryPattern, BMerge, BKnit, ScopeOptions)` [5] [DOC]
- `HoleWizard5(GenericHoleType, StandardIndex, FastenerTypeIndex, SSize, EndType, Diameter, Depth, Length, Value1..Value12, ThreadClass, RevDir, FeatureScope, AutoSelect, AssemblyFeatureScope, AutoSelectComponents, PropagateFeatureToParts)` [27] [DOC]
- `CreateDefinition(swFmXxx)` → set properties → `CreateFeature(data)`: the feature-data route (sweep [PROVEN])
- `InsertCombineFeature(long OperationType, IBody2 MainBody, variant ToolVar)` [3] → IFeature. For Add: `(SWBODYADD, NOTHING, dispatch_array(bodies))` [PROVEN]
- `InsertFeatureTreeFolder2(Type)` [1]

**IAssemblyDoc**
- `AddComponent5(CompName, ConfigOption, NewConfigName, UseConfigForPartReferences, ExistingConfigName, X, Y, Z)` [8]
- `AddComponents3(Names, Transforms, CoordinateSystemNames)` [3] (batch insert) [DOC]
- `AddMate5(MateTypeFromEnum, AlignFromEnum, Flip, Distance, DistanceAbsUpperLimit, DistanceAbsLowerLimit, GearRatioNumerator, GearRatioDenominator, Angle, AngleAbsUpperLimit, AngleAbsLowerLimit, ForPositioningOnly, LockRotation, WidthMateOption, ErrorStatus)` [15]
- `CreateMateData(swMateType)` → `CreateMate(MateData)`: the current official route (section 10.4)
- `FixComponent()` / `UnfixComponent()` act on the selected components

**Configurations and suppression** [PROVEN]
- `IModelDoc2.AddConfiguration3(Name, Comment, AlternateName, Options)` [4] → object (None = failed)
- `IModelDoc2.ShowConfiguration2(ConfigurationName)` [1] (activates it) · `GetConfigurationNames()` [0]
- `IModelDoc2.EditSuppress2()` / `EditUnsuppress2()` [0]: act on the selected features in the **active** configuration
- `IFeature.IsSuppressed2(Config_opt, Config_names)` [2] → array of bool: `IsSuppressed2(swThisConfiguration, None)[0]`
- `IFeature.SetSuppression2(SuppressionState, Config_opt, Config_names)` [3]: with swSpecifyConfiguration it returned True and did nothing (7.11)

**Measuring** [PROVEN]
- `IBody2.GetExtremePoint(X, Y, Z, ByRef Outx, ByRef Outy, ByRef Outz)` [6] → `(ok, x, y, z)`: the exact furthest point along a direction
- `IFace2.GetClosestPointOn(X, Y, Z)` [3] → `(x, y, z, ...)` or None
- `ICurve.Evaluate2(Parameter, NumberOfDerivatives)` [2] → point, then derivatives: `[3:6]` is the tangent
- `IFeature.GetErrorCode2(ByRef IsWarning)` [1] → error code (0 = OK). Early-bound: `code, warn = f.GetErrorCode2()`

**Others**: `IEquationMgr.Add2(Index, Equation, Solve)` / `Add3(..., WhichConfigurations, ConfigNames)`,
`IEquationMgr.Equation(Index)` (indexed property: get with `eq.Equation(i)`, put with
`com_put(eq, "Equation", i, text)`; early-bound makepy also has `eq.SetEquation(i, text)`),
`IEquationMgr.GetCount()`, `IEquationMgr.EvaluateAll()`,
`IPartDoc.SetMaterialPropertyName2(Config, Database, Name)` (fails silently, see 7.10),
`IPartDoc.GetMaterialPropertyName2(ConfigName, ByRef Database)` → name,
`IPartDoc.GetBodies2(BodyType, VisibleOnly)`, `IPartDoc/IAssemblyDoc/IComponent2.FeatureByName(Name)`,
`IComponent2.Select4(Append, Data, ShowPopup)`, `IComponent2.ReferencedConfiguration` (get/put),
`IComponent2.Solving` (makepy: get only, so set it with `com_put`), `IFeature.Select2(Append, Mark)`,
`ISldWorks.GetOpenDocumentByName(DocumentName)`.

---

## 7. Part modelling recipes

### 7.1 New part and default planes [PROVEN]

```python
doc = cast(app.NewDocument(app.GetUserPreferenceStringValue(SWC.swDefaultTemplatePart), 0, 0, 0), "IModelDoc2")
if doc is None: raise SystemExit("No default part template - check Tools > Options > Default Templates")
ext = cast(doc.Extension, "IModelDocExtension")
ext.SetUserPreferenceInteger(SWC.swUnitSystem, 0, SWC.swUnitSystem_MMGS)
sm = cast(doc.SketchManager, "ISketchManager"); fm = cast(doc.FeatureManager, "IFeatureManager")
sm.AddToDB = True; sm.DisplayWhenAdded = False        # exact geometry, no inferencing, faster

# The first three RefPlane features are always Front, Top, Right in any language edition.
planes, f = [], doc.FirstFeature()
while f is not None and len(planes) < 3:
    f = cast(f, "IFeature")
    if f.GetTypeName2() == "RefPlane": planes.append(f)
    f = f.GetNextFeature()
front, top, right = planes          # normals: Front = +Z, Top = +Y, Right = +X
```

Localised editions name the planes differently, e.g. Chinese 前视基准面 (Front), 上视基准面 (Top),
右视基准面 (Right). Finding them by type order works in every language. If you must use names,
English is `"Front Plane"`, `"Top Plane"`, `"Right Plane"`, selected with type `"PLANE"`.

### 7.2 Builder preferences: turn off, then restore [PROVEN]

```python
saved = {t: app.GetUserPreferenceToggle(t) for t in (SWC.swInputDimValOnCreate, SWC.swSketchInference)}
for t in saved: app.SetUserPreferenceToggle(t, False)
try:
    build()
finally:
    for t, v in saved.items(): app.SetUserPreferenceToggle(t, v)
```

- `swInputDimValOnCreate` off: `AddDimension2` doesn't open the "Modify" pop-up, which would hang
  the script.
- `swSketchInference` off: without it, `CreateCenterRectangle` snapped an inner rectangle onto the
  outer one even with `AddToDB = True`.

### 7.3 Sketch coordinates vs model coordinates [PROVEN]

Sketch points are in the **sketch's own 2D frame**, which is not the model frame:
- Front Plane: sketch (u, v) = model (X, Y).
- Top Plane: sketch (u, v) = model (X, **−Z**), as measured in the enclosure build.
- Right Plane and offset planes: don't assume. Use the sketch transform:

This version needs no `IMathUtility`, so it works early- and late-bound:

```python
a = cast(sm.ActiveSketch, "ISketch").ModelToSketchTransform.ArrayData   # 16 doubles
# SOLIDWORKS transforms use row vectors: p' = scale * (p . R) + t
# a[0:9] = R (rows), a[9:12] = t, a[12] = scale
def to_sketch(x, y, z):                       # model mm -> sketch (u, v) metres
    p = (x*MM, y*MM, z*MM)
    return tuple(a[12] * sum(p[i] * a[3*i + j] for i in range(3)) + a[9 + j] for j in range(2))
```

Early-bound only, the `IMathUtility` route also works:
`cast(mu.CreatePoint(darr([x*MM, y*MM, z*MM])), "IMathPoint").MultiplyTransform(xf)`. Late-bound,
it throws *The server threw an exception*. To *set* a transform (`IComponent2.Transform2`), you
need `CreateTransform`, so get an early-bound `IMathUtility`:
`SW.IMathUtility(app.GetMathUtility()._oleobj_.QueryInterface(SW.IMathUtility.CLSID, pythoncom.IID_IDispatch))`.

The same transform tells you the sketch normal: transform (0,0,0) and the model unit axis and
compare the z values (row vector: `z = sum(p[i] * a[3*i + 2])`). That sign gives the default
extrude direction (7.5). In general, the sketch's +Z in model space is the column `R[:, 2]`; set
`Dir` from `sign(sketch_normal · wanted_direction)`. This works on any plane, including oblique
ones [PROVEN]. When you draw arcs in your own (u, a) frame, flip the arc direction if that frame
is left-handed in sketch space.

### 7.4 Reference planes: offset, verified [PROVEN]

```python
def plane_origin(feat):
    """Origin (mm) and unit normal of a reference plane feature."""
    rp = cast(cast(feat, "IFeature").GetSpecificFeature2(), "IRefPlane")
    a = cast(rp.Transform, "IMathTransform").ArrayData
    return [v / MM for v in a[9:12]], [a[6], a[7], a[8]]

def offset_plane(base_feat, axis_index, value_mm, name):
    """Plane parallel to base_feat at signed coordinate value_mm on model axis 0/1/2.
    The side SOLIDWORKS picks isn't guaranteed: read it back, delete and retry flipped."""
    flags = SWC.swRefPlaneReferenceConstraint_Distance
    for _ in range(2):
        doc.ClearSelection2(True); base_feat.Select2(False, 0)
        if fm.InsertRefPlane(flags, abs(value_mm) * MM, 0, 0, 0, 0) is None: raise RuntimeError(name)
        f = cast(doc.FeatureByPositionReverse(0), "IFeature")      # part only (rule 1.6)
        if abs(plane_origin(f)[0][axis_index] - value_mm) < 1e-6:
            f.Name = name
            return f
        doc.ClearSelection2(True); f.Select2(False, 0); ext.DeleteSelection2(0)
        flags ^= SWC.swRefPlaneReferenceConstraint_OptionFlip
    raise RuntimeError(f"could not place {name} at {value_mm}")
```

The side an offset plane lands on is **not guaranteed**. Always read back
`Transform.ArrayData[9:12]` (the origin) and retry with `OptionFlip` toggled if it's wrong.
Cache planes by (axis, value) so you don't create duplicates.

**Angle planes** (`InsertRefPlane(swRefPlaneReferenceConstraint_Angle, rad, swRefPlaneReferenceConstraint_Coincident, 0, 0, 0)`
with a plane and an axis selected): the normal can come out reversed, as R(−θ) or R(θ)+180°.
Read the normal back (`ArrayData[6:9]`). If it is off by 180°, keep the plane and swap the
alignment of any mate that uses it. If it is off by anything else, delete it and retry with
`OptionFlip`.

**Plane through any 3 model points** [PROVEN] (an oblique wall, a tilted skin):

```python
sm.Insert3DSketch(True)
pts = [sm.CreatePoint(x*MM, y*MM, z*MM) for x, y, z in (p1, p2, p3)]   # model coordinates
sm.Insert3DSketch(True)                                                # close the 3D sketch
doc.ClearSelection2(True)
for i, pt in enumerate(pts):
    sd = cast(selm.CreateSelectData(), "ISelectData"); sd.Mark = i      # marks 0, 1, 2
    cast(pt, "ISketchPoint").Select4(i > 0, sd)
c = SWC.swRefPlaneReferenceConstraint_Coincident
fm.InsertRefPlane(c, 0, c, 0, c, 0)
```

Read back `IRefPlane.Transform` (origin `[9:12]`, normal `[6:9]`). It matched to < 0.001 mm. Hide
the helper 3D sketch afterwards (`BlankSketch`).

### 7.5 Extrude and cut, with the right direction [PROVEN]

Measured on SOLIDWORKS 2025: a **boss** goes along +sketch normal by default and a **cut** goes
along −sketch normal by default. `Dir=True` reverses the direction. `Flip` flips the side to cut
and must stay `False` for normal cuts.

```python
# boss, blind, merge into the body
f = fm.FeatureExtrusion3(True, False, reverse, SWC.swEndCondBlind, 0, depth*MM, 0,
                         draft, False, draft_outward, False, math.radians(draft_deg), 0,
                         False, False, False, False, True, True, True, 0, 0, False)
# cut, blind (through_all: T1 = swEndCondThroughAll)
f = fm.FeatureCut4(True, False, reverse, SWC.swEndCondBlind, 0, depth*MM, 0,
                   draft, False, draft_outward, False, math.radians(draft_deg), 0,
                   False, False, False, False, False, True, True, False, False, False, 0, 0, False, False)
# symmetric through-all-both cut
f = fm.FeatureCut4(False, False, False, SWC.swEndCondThroughAll, SWC.swEndCondThroughAll,
                   0.01, 0.01, False, False, False, False, 0, 0, False, False, False, False,
                   False, True, True, False, False, False, 0, 0, False, False)
```

Select the sketch feature first (`sketch.Select2(False, 0)` after `ClearSelection2`). Belt and
braces: extrude, then check that the volume changed in the expected sense, that there is one body,
and where the far face is. If the check fails, delete the feature (keep its sketch) and retry with
`reverse` toggled.

Check only the **far end** of a new feature's faces. Coplanar faces merge into older features,
so the near face may belong to an earlier feature.

### 7.6 Revolve, loft, sweep, helix [PROVEN]

```python
# Revolve the selected sketch about its single centerline (360 deg)
fm.FeatureRevolve2(True, True, False, is_cut, False, False, SWC.swEndCondBlind, 0,
                   math.radians(360), 0, False, False, 0, 0, 0, 0, 0, True, True, True)
# Loft boss: select profile sketches in order with Mark 1
a.Select2(True, 1); b.Select2(True, 1)
fm.InsertProtrusionBlend2(False, True, False, 1.0, 0, 0, 1.0, 1.0, True, True,
                          False, 0.0, 0.0, 0, True, True, True, 0)
# Loft cut
fm.InsertCutBlend(False, True, False, 1.0, 0, 0, False, 0, 0, 0, True, True)
# Helix on a selected circle sketch, then circular-profile sweep along it
doc.InsertHelix(False, left_hand, False, False, SWC.swHelixDefinedByHeightAndRevolution,
                height*MM, pitch*MM, revolutions, 0, 0)       # Clockwised=True -> left-hand (below)
helix.Select2(False, 4)                               # path = Mark 4
sd = fm.CreateDefinition(SWC.swFmSweep)
sd.CircularProfile = True; sd.CircularProfileDiameter = wire_d * MM
sd.Merge = True                                       # False -> a separate body (7.11)
fm.CreateFeature(sd)                                  # None = failed
```

- For a loft, use **separate sketches** for the outer and inner loops: loft the outer ones as a
  boss, then the inner ones as a loft cut.
- **Oblique prism cut** [PROVEN] (a slanted slot, a chevron leg): loft cut
  `InsertCutBlend(False, True, False, 1.0, 0, 0, False, 0, 0, 0, True, True)` between two identical
  profiles on parallel planes, offset sideways. Run each leg 0.3 mm past its start and end planes
  (along the same slope), so it cuts cleanly through. 56 of them on a 166-feature part rebuilt in
  ~4 min with graphics off. On flat walls, share the planes within a group. Use per-slot planes only
  on curved skins.
- A single sweep along a joined helix with pitch changes fails, because the joins are slight
  corners. Sweep each helix separately; they merge into one body. Check `len(GetBodies2) == 1`.

**Helix hand and start** [PROVEN, SW 2025, circle sketched on a plane parallel to Top, normal +Y]:
- `InsertHelix(Reversed=False, Clockwised=True, ...)` gives a **left-hand** helix advancing **+Y**.
  `Clockwised=False` gives right-hand.
- It starts at the sketch plane, on the circle at model **+Z**, and a left-hand helix leaves the
  start travelling **−X**.
- Feature dimensions: **D3 = height, D5 = revolutions**. Link those in equations.
- Always check the hand from the curve itself, and stop the build if it is wrong:

```python
edge = cast(cast(helix, "IFeature").GetSpecificFeature2(), "IReferenceCurve").GetFirstSegment()
cp = cast(cast(edge, "IEdge").GetCurveParams3(), "ICurveParamData")
s, e = [v / MM for v in cp.StartPoint], [v / MM for v in cp.EndPoint]
t = cast(edge.GetCurve(), "ICurve").Evaluate2(cp.UMinValue, 1)[3:6]     # tangent at the start
hand = "RH" if (s[2] * t[0] - s[0] * t[2]) * (e[1] - s[1]) > 0 else "LH"   # axis = +Y
```

**A feature that only touches the body won't merge** [PROVEN]. Face to face (a sweep starting
exactly on the coil wire's end face), edge to face, or tangent (a leg grazing the next coil
turn): the result is `None`, or 2 bodies.
- Make the new feature **overlap** the body by 0.2–0.5 mm (start a leg path 0.3 mm back inside the
  wire), or give it a **clear gap** of at least 0.01 mm from faces it mustn't touch (the leg was
  moved 0.02 mm axially away from the next turn).
- To debug a failed merge, re-run the feature with `Merge = False` and print each body's
  `GetBodyBox()`. That shows where the bodies meet.
- To join bodies later: `fm.InsertCombineFeature(SWC.SWBODYADD, NOTHING, dispatch_array(bodies))`,
  then check that there is 1 body and What's Wrong = 0.

### 7.7 Shell, fillet, chamfer: select by geometry [PROVEN]

```python
# shell: select the faces to remove with Mark 1 via ISelectData
sd = cast(selm.CreateSelectData(), "ISelectData"); sd.Mark = 1
for i, face in enumerate(faces): cast(face, "IEntity").Select4(i > 0, sd)
doc.InsertFeatureShell(t*MM, False)                  # check last feature type == "Shell"
# fillet edges or faces (Options 195 = value proven in our builds)
fm.FeatureFillet3(195, r*MM, 0, 0, 0, 0, 0, None, None, None, None, None, None, None)
# chamfer, angle-distance (Options 4 = value proven in our builds)
fm.InsertFeatureChamfer(4, SWC.swChamferAngleDistance, d*MM, math.radians(45), 0, 0, 0, 0)
```

Fillet the outer perimeter **before** shelling. That keeps walls uniform and the shell succeeds
more often. Check that a chamfer really cut something: a chamfer feature once came out with no
faces and no volume change. Assert that `f.GetFaces()` isn't empty or that the volume dropped.

### 7.8 Global variables and fully defined sketches [PROVEN]

```python
eq = cast(doc.GetEquationMgr(), "IEquationMgr")
eq.Add2(-1, '"WALL" = 2', True)                                  # global variable
eq.Add2(-1, '"TIP_W" = "BASE_W" + 2 * "LEN" * tan ( "FOV_H" )', True)   # derived
# dimension a sketch entity, then drive it
seg.Select4(False, None); dd = doc.AddDimension2(tx*MM, ty*MM, 0)
dim_name = dd.GetDimension2(0).Name                              # e.g. "D1"
# after closing and naming the sketch:
eq.Add2(-1, f'"{dim_name}@Flange Sketch" = "FLANGE_W"', True)
# drive a feature dimension (extrude depth, plane offset) - D1, D2 ... in creation order
dd = feat.GetFirstDisplayDimension()  # walk with feat.GetNextDisplayDimension(dd)
eq.Add2(-1, '"D1@Flange" = "FLANGE_THK"', True)
```

- Two ways to fully define a sketch: (a) dimensions linked to variables plus relations
  (`sgATMIDDLE` to centre a rectangle diagonal on `Point1@Origin`, `sgCOINCIDENT` for a hole
  centre on a construction corner), or (b) quick: select all segments and `sgFIXED`.
- Check: `feat.GetSpecificFeature2().GetConstrainedStatus() == 3` for every sketch.
- Test the parametrics: change 3–4 variables, `ForceRebuild3`, check What's Wrong = 0, then
  restore the values.
- **Changing an existing equation** [PROVEN]: delete followed by `Add2` failed. Set it in place.
  `Equation(i)` is an *indexed property*, so use `com_put` (3.0):

```python
def set_global(eq, name, value):
    for i in range(eq.GetCount()):               # late-bound: call(eq, "GetCount")
        if eq.Equation(i).replace(" ", "").startswith(f'"{name}"='):
            com_put(eq, "Equation", i, f'"{name}" = {value:g}')     # early-bound also: eq.SetEquation(i, ...)
            if eq.Equation(i).replace(" ", "") != f'"{name}"={value:g}':
                raise RuntimeError(f"could not set {name}: {eq.Equation(i)}")
            break
    else:
        raise RuntimeError(f"no global variable {name}")
eq.EvaluateAll(); doc.ForceRebuild3(False)
```
- The origin selects as `ext.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, True, 0, None, 0)`.

### 7.9 Text and labels [PROVEN]

- `doc.InsertSketchText(x, y, 0, "TEXT", 0, 0, mirror, width_factor_pct, spacing_pct)`, then set
  `ITextFormat` (`CharHeight` in metres, `TypeFaceName`, `Bold`, `Escapement` = rotation in
  radians) and call `st.SetTextFormat(False, tf)`.
- Text on a side wall reads mirrored or rotated depending on the plane. For dependable embossed or
  engraved labels, draw the glyph outlines as sketch lines (matplotlib `TextPath` → polygons).
  This works on any plane. See `swlib.text_outline`.

### 7.10 Appearance, material, properties [PROVEN]

`MaterialPropertyValues` is 9 doubles:
**[R, G, B, Ambient, Diffuse, Specular, Shininess, Transparency, Emission]**, all 0–1.
Index 7 is transparency. Putting a value there by mistake makes the part see-through, which is a
bug we hit.

```python
doc.MaterialPropertyValues = darr([r, g, b, 1.0, 0.8, 0.3, 0.3, 0.0, 0.0])        # ACTIVE configuration
cast(face, "IFace2").MaterialPropertyValues = darr([...])                          # one face (accent)
part.SetMaterialPropertyName2("", "SOLIDWORKS Materials", "PET")                     # "" = active config
cpm = cast(ext.CustomPropertyManager(""), "ICustomPropertyManager")
cpm.Add3("Description", SWC.swCustomInfoText, "Front shell", SWC.swCustomPropertyReplaceValue)
```

- **Appearance is per configuration** [PROVEN]. A colour set in Default doesn't show in the other
  configurations, so their components were grey in the assemblies. Activate each configuration
  (`ShowConfiguration2`) and set `MaterialPropertyValues` in each.
- **`SetMaterialPropertyName2` fails silently** when the name isn't in the database. AISI 302 isn't
  in `SOLIDWORKS Materials.sldmat`. Always read the material back and fall back to an
  equivalent (AISI 304 has the same density), and say so in the report. Missing grades seen so
  far, with the substitute: AISI 302 / 301 / 303 → *AISI 304*, AISI 420 → *Chrome Stainless Steel*
  (14.2):

```python
part.SetMaterialPropertyName2("", "SOLIDWORKS Materials.sldmat", "AISI 302")
name, db = part.GetMaterialPropertyName2("")                     # early-bound: Database is a ByRef output
# late-bound: db = VARIANT(VT_BYREF | VT_BSTR, ""); name = part.GetMaterialPropertyName2("", db)
if not name:
    part.SetMaterialPropertyName2("", "SOLIDWORKS Materials.sldmat", "AISI 304")
```

### 7.11 Configurations [PROVEN]

Used for rigid pieces of a flexible sub-assembly (piece A = coil + leg 1, piece B = leg 2), size
variants and simplified versions.

**Plan the feature tree first.** A feature that must exist on its own in a configuration must not
be a child of a feature that is suppressed there. Suppressing a parent suppresses its children.
A leg merged into the coil is the coil's child, so it vanished with the coil. Build that feature
as a **separate body** (`Merge = False`), and join it with a Combine feature at the end of the
tree ("Join Leg 2"). The pieces then suppress the join, not the feature.

```python
for cfg in ("Piece A", "Piece B"):
    if not doc.AddConfiguration3(cfg, "rigid piece", "", 0): raise RuntimeError(cfg)

def suppress_in(cfg, names):
    doc.ShowConfiguration2(cfg)                      # activate: EditSuppress2 acts on the active config
    doc.ClearSelection2(True)
    for i, n in enumerate(names):                    # look features up AGAIN by name after a switch
        cast(part.FeatureByName(n), "IFeature").Select2(i > 0, 0)
    doc.EditSuppress2(); doc.ClearSelection2(True); doc.ForceRebuild3(False)

suppress_in("Piece A", ("Leg 2 Wire", "Join Leg 2"))
suppress_in("Piece B", ("Join Leg 2", "Leg 1 Wire", "Coil Wire"))
# read back every configuration
for cfg, want in expect.items():
    doc.ShowConfiguration2(cfg)
    got = tuple(bool(cast(part.FeatureByName(n), "IFeature").IsSuppressed2(SWC.swThisConfiguration, None)[0])
                for n in names)
    if got != want: raise RuntimeError(f"{cfg}: {got}, wanted {want}")
```

- `SetSuppression2(state, swSpecifyConfiguration, names)` returned True and suppressed nothing.
  Use the activate, select, `EditSuppress2` route above.
- **Feature handles go stale after a configuration switch.** Never reuse an `IFeature` from before
  `ShowConfiguration2`. Look it up again by name.
- **A temporary suppress also suppresses dependents.** After suppressing the legs to measure the
  coil alone, unsuppressing only the legs left "Join Leg 2" suppressed, and Default had 2 bodies.
  Unsuppress the whole list, *including dependents*.
- **Just before saving**, check every configuration: body count, volume, What's Wrong = 0, and the
  suppression table. Then activate the configuration the user should see first.

### 7.12 Big sketches, engraving, and naming entities for later [PROVEN]

- **No more than ~40 segments in one sketch.** 201 engraved scale lines plus 21 digits in one sketch,
  all fixed with `fix_all`, took SOLIDWORKS ~20 min to solve. Draw **one** line and use a linear
  pattern, or split into sketches of ~20 lines. Time a 10-line test run first, then extrapolate.
- **Engrave last, behind a configuration.** Engraved scales and text give a part thousands of small
  faces (a beam had ~5,000), and every later per-face or per-edge COM walk becomes slow. Put the
  marking features last, and add a `NO_MARKING` configuration with them suppressed. Drawing views
  and edge searches use that configuration; show the engraving only in a detail view.
- **Name the faces and edges you will dimension or mate to, while you build the part**, so a drawing
  or an assembly script finds them directly instead of searching:
  `part.SetEntityName(face, "BEAM_TOP")`, later `part.GetEntityByName("BEAM_TOP", SWC.swSelFACES)`
  [PROVEN 2026-10-31, `swlib.Part.name_entity` / `entity`]. Without names, read only the **outer loops of faces owned by the non-marking
  features** (`IFeature.GetFaces()`, `IFace2.GetLoops()`, `ILoop2.IsOuter()`) and cache them once per
  part [PROVEN]. That shortcut missed edges on **inner** loops (a window, a channel), so fall back to
  a full search for those.
- **Turned parts** (screws, pins, rivets) have no straight outline edges. Find them by **circular
  edges**: a diameter from a circle, a length between two circles seen edge-on.

Colour policy (user preference): **never model in black**, because renders become unreadable.
Use white/grey/colour per part, with accents on named feature faces. Recolour with a separate
`recolor.py`. **Never rebuild geometry just to change colours.**

---

## 8. Saving, exporting, rendering [PROVEN]

```python
ok = ext.SaveAs3(path_sldprt, 0, SWC.swSaveAsOptions_Silent, None, None, 0, 0)            # save
ok = ext.SaveAs3(path_step,   0, SWC.swSaveAsOptions_Silent | SWC.swSaveAsOptions_Copy, None, None, 0, 0)  # export
# the extension decides the format: .STEP .STL .IGS .x_t .jpg ...
# IModelDoc2.SaveAs3(name, 0, options) [3 args] returns an int error code (0 = OK)
# Extension.SaveAs3 early-bound returns (ok, errors, warnings); decode errors with swFileSaveError_e:
# 2049 = swGenericSaveError (1) | swFileSaveAsNameExceedsMaxPathLength (2048) -> path too long [PROVEN]

app.SetUserPreferenceIntegerValue(SWC.swStepAP, 214)                       # STEP AP214 (colours)
app.SetUserPreferenceToggle(SWC.swSTLBinaryFormat, True)
app.SetUserPreferenceIntegerValue(SWC.swExportStlUnits, SWC.swMM)
app.SetUserPreferenceIntegerValue(SWC.swSTLQuality, SWC.swSTLQuality_Custom)
app.SetUserPreferenceDoubleValue(SWC.swSTLDeviation, 0.01 * MM)
app.SetUserPreferenceDoubleValue(SWC.swSTLAngleTolerance, math.radians(5))
app.SetUserPreferenceToggle(SWC.swSTLDontTranslateToPositive, True)        # keep model coords
```

Keep output paths well under 260 characters. A save into a deep temp folder (about 260
characters) failed with 2049.

**Save over an earlier result: close it by path, never by its old title** [PROVEN]. After `SaveAs3`
the document's title changes, so `CloseDoc(old_title)` closes nothing and the next save fails
(error 1: the previous run's file is still open). Before saving, close any open copy of the target
path (`swlib.SwApp.close_path(path)`). After `SaveAs3`, read the title again.

**Author and file properties** (useful for drawings' DRAWN field and for traceability):
`com_put(doc, "SummaryInfo", SWC.swSumInfoAuthor, name)` (an indexed property; read back with
`doc.SummaryInfo(SWC.swSumInfoAuthor)`) [PROVEN 2026-10-31], and a custom property `Author`.
Exported files [PROVEN 2026-10-31]: PDF with `pymupdf` `set_metadata({"author": ..., "creator": ...})`
then `saveIncr()`; STEP by rewriting the author list in the header `FILE_NAME(name, time, (author),
...)` (read and write latin-1, double any `'`). `swlib.stamp_pdf` / `stamp_step` do both.

**Hidden data inside a model: a hidden attribute** [PROVEN 2026-10-31]. It is saved with the file,
isn't shown in the FeatureManager tree, survived save, close and reopen, and appears in a feature
traversal as type `Attribute`:

```python
d = cast(app.DefineAttribute("MyDataDef"), "IAttributeDef")
for p in ("key", "value"):
    d.AddParameter(p, SWC.swParamTypeString, 0.0, 0)
d.Register()                         # False if this SOLIDWORKS session registered the name already
att = cast(d.CreateInstance5(doc, None, "MyData", 1, SWC.swAllConfiguration), "IAttribute")  # 1 = hidden
cast(att.GetParameter("key"), "IParameter").SetStringValue2("abc", SWC.swAllConfiguration, "")
# read back: f.GetTypeName2() == "Attribute" and f.Name.startswith("MyData") in features(doc)
val = cast(cast(f.GetSpecificFeature2(), "IAttribute").GetParameter("key"), "IParameter").GetStringValue()
```

- The definition belongs to the SOLIDWORKS session, not to the script. **Write and verify in the same
  process.** A second process that defined the same attribute again and opened the stamped files was
  followed by a SOLIDWORKS crash (fault module `slduiu`). The cause isn't proven; until a Test Run
  separates the cases, check from a second process only on request, with the user's work saved.
- `IModelDoc2.IGet3rdPartyStorage` returned None for write and read from an external script. It needs
  an add-in's save notification. From a script, use the attribute above, or at file level an NTFS
  alternate data stream: `open(f"{path}:name", "w")`. The stream is invisible in Explorer but lost when
  the file is zipped, uploaded or copied to a non-NTFS drive, so it can only be a second copy.

### 8.1 Renders [PROVEN]

`ShowNamedView2("*Isometric", 7)` (or `"*Front"`, `-1`), `ViewZoomtofit2()`, `GraphicsRedraw2()`,
`SaveBMP(path.bmp, 1600, 1200)`, then convert to PNG with Pillow. Hide reference geometry first with
`SetUserPreferenceToggle(swViewDisplayHideAllTypes, True)`. Turn graphics updates back on before
rendering.

- `ViewZoomTo2` did nothing in a part window, and zoom-fit clipped top and bottom views. Use
  `ViewZoomtofit2()`, then `IModelView.Scale2 *= 0.7`, then `SaveBMP(f, 3832, 3288)`.
- **Exact model-to-pixel mapping** (to draw overlays such as slots, arrows and keep-outs onto a
  render): `a = IModelView.Transform.ArrayData`, `pix = s * (p @ R)[:2] + t` (row vectors). Scale it by
  the saved size divided by the `SaveBMP(f, 0, 0)` size, **separately for x and y**.
  `SaveBMP(0, 0)` size is not `FrameWidth`. Crop with this transform, not by auto-detecting the
  background: a gradient background broke auto-crop.
- X-ray views of an assembly: open it read-only, set `comp.Visible = 0/1`, render, close it.
- Report sheets: use one tested template, ~3000–3400 px wide, labels 50–58 px, text wrapped at ~100
  characters, plain glyphs only (Segoe UI has no ⟂ ≫ ✓). 7000 px sheets came out clipped and
  overlapping. Multi-page PDF: matplotlib `PdfPages`, one page per part. Preview pages as PNG with
  `pymupdf`. To read a PDF from the web (a datasheet), download it and use `pymupdf` `get_text()`.

### 8.2 Drawings through the API [PROVEN 2026-10-30, a 7-sheet caliper drawing set]

- **Sheet**: `NewDocument(drawing_template, swDwgPaperA3size, 0.42, 0.297)`, then
  `SetupSheet6(Name, PaperSize, TemplateIn, Scale1, Scale2, FirstAngle, TemplateName, Width, Height,
  PropertyViewName, RemoveModifiedNotes, ZoneLeft, ZoneRight, ZoneTop, ZoneBottom, ZoneRow, ZoneCol)` [17]
  with `swDwgPaperA3size`, `swDwgTemplateCustom`, `FirstAngle=True` (ISO) and
  `"<lang>\sheetformat\a3 - iso.slddrt"`. After the views, call `ISheet.SetScale(n, d, False, False)`,
  or the title block shows the wrong scale.
- **Views**: `CreateDrawViewFromModelView3(path, "*Front", x, y, 0)` → `IView`; set `ScaleDecimal`;
  `SetDisplayMode4(False, swHIDDEN | swHIDDEN_GREYED, ...)`. Model → sheet point:
  `IView.ModelToViewTransform`. For a part with engraving, reference the `NO_MARKING` configuration (7.12).
- **Dimensions**: `IView.SelectEntity(model_edge, append)` takes the **part's own** `IEdge`, then
  `AddDimension2(x, y, 0)` in sheet metres. Two parallel line edges give a distance; a circle gives a
  diameter. **Never overwrite dimension text** to fake a value (a radius dimension retyped "Ø7.0"
  no longer follows the model). Use a true diameter dimension; add only a prefix or suffix.
- **Tolerances**: `IDimension.SetToleranceType(swTolBILAT | swTolSYMMETRIC | swTolFITWITHTOL)` and
  `IDimensionTolerance.SetValues(min, max)` (metres). `SetValues2` returned False on drawing dimensions.
  Fits: `SetFitValues("H7", "")` or `("", "g6")`. Text: `IDisplayDimension.SetText(swDimensionTextAll, ...)`.
- **GD&T** (2023+ format): `doc.InsertGtol()` with the edge selected, then
  `SetFrameValues2(1, "0.005", "", "A", "", "")`, `SetFrameSymbols2(1, "<IGTOL-PERP>", ...)`, then
  **`ConvertFormat()`**, or the symbol isn't drawn. Symbol names come from `gtol.sym`: `<IGTOL-PERP>`,
  `<IGTOL-FLAT>`, `<IGTOL-PARA>` (**not** `PARALLEL`, which printed as raw text), `<IGTOL-STRAIGHT>`.
  Look up any other name in `gtol.sym` before using it.
- **Datums**: `InsertDatumTag2()` + `SetLabel("A")`. **Notes**: `InsertNote(text)`, with `chr(10)` for a
  new line (not `<br>`). `IAnnotation.Color = 0` prints black (dimensions added by the API are grey
  otherwise).
- **Surface finish**: `Extension.InsertSurfaceFinishSymbol3(SymType, LeaderType, X, Y, Z, LaySymbol,
  ArrowType, MachAllowance, OtherVals, ProdMethod, SampleLen, MaxRoughness, MinRoughness,
  RoughnessSpacing)` [14]. With the ISO template, ProdMethod printed and the roughness arguments did
  not. **Don't** put "Ra 0.1" in the method slot (that is non-standard). Find the argument that prints
  in the requirement position with a test run, and set the annotation text height (it is large).
- **Auto-arrange**: try `Extension.AlignDimensions(swAlignDimensionType_AutoArrange, spacing)` as a
  first tidy-up pass. [DOC: test run first]
- **Dry-run every drawing first**: resolve every edge you will dimension and report *all* the missing
  ones in one pass, before creating any sheet.

### 8.3 Drawing layout and the GD&T checklist

Six reruns went on one problem at a time: a 4:1 end view off the sheet, huge surface-finish text,
notes past the border, a BOM off the sheet and then over the title block, and texts on texts.
Plan the layout instead:
1. **Lay out the sheet before writing code.** Use fixed zones: a view grid, a notes zone, a BOM zone
   above the title block, and a 10 mm margin round each view. Choose view scales so that each view
   *plus its dimensions* fits its zone (an end view's height includes the jaws).
2. **ISO 129 spacing**: the first dimension line ≥ 10 mm from the outline, then ≥ 7 mm between parallel
   dimension lines. Stagger stacked texts, and put shorter dimensions inside longer ones.
3. **GD&T frames, datum symbols and finish symbols go outside the view**, never over engraving,
   digits or hidden lines. A datum must sit on a **visible** outline or extension line, never on a
   hidden (dashed) line. Use a section or end view.
4. **Small features get a detail view** (jaw tips, a relief groove, a graduation pattern) instead of
   crowding a 1:1 view.
5. **Check every PDF for collisions automatically**: read text boxes and line boxes with `pymupdf`
   (`page.get_text("dict")`, `page.get_drawings()`). Flag text on text, and text crossing a view
   outline. Move the offender and re-run only that sheet.
6. **GD&T checklist** the drawing script must pass:
   - Ⓔ (envelope) on every feature-of-size fit when the sheet declares ISO 8015, or a note "ISO 14405 Ⓔ
     applies to all fits".
   - Every *located* feature (a groove, a hole position, an edge) has **basic (boxed) dimensions plus a
     position or profile frame** to the datums, never ± location dimensions, and never only a note.
   - A coplanarity or flatness requirement is a frame (flatness, or profile of a surface to a datum),
     not a note. A straightness on a long edge has a unit length ("0.003/150").
   - Surface texture per ISO 21920 / ISO 1302 (Ra in the requirement position, the method as the
     method), plus a general surface-texture symbol near the title block.
   - Title block filled: TOLERANCES / LINEAR / ANGULAR (e.g. ISO 2768-mK), CHK'D, WEIGHT, material
     (state the real grade when the library had none, 7.10), projection symbol.
   - An assembly drawing has at least its overall dimensions and range of motion.

---

## 9. Selecting faces and edges by geometry [PROVEN]

```python
def faces(part):
    for b in part.GetBodies2(SWC.swSolidBody, True) or ():
        for f in cast(b, "IBody2").GetFaces() or (): yield cast(f, "IFace2")

def find_faces(part, normal, axis, at, tol=1e-3):
    """Planar faces with unit `normal` lying at coordinate `at` (mm) on axis 0/1/2."""
    out = []
    for f in faces(part):
        if not cast(f.GetSurface(), "ISurface").IsPlane(): continue
        if sum((a - b) ** 2 for a, b in zip(f.Normal, normal)) > 1e-6: continue
        box = [v / MM for v in f.GetBox()]
        if abs(box[axis] - at) < tol and abs(box[axis + 3] - at) < tol: out.append(f)
    return out
```

- Edges: `face.GetEdges()` gives `IEdge`. `GetCurveParams3()` gives start/end points.
  `GetCurve().IsCircle()` and `.CircleParams` find hole edges by centre.
- A feature's faces: `feat.GetFaces()`. Use this for per-feature colour accents and to check a
  feature's extent.
- `SelectByID2` by name: planes `("Front Plane", "PLANE")`, sketches `("Sketch1", "SKETCH")`,
  axes `("Axis1", "AXIS")`, origin `("Point1@Origin", "EXTSKETCHPOINT")`. Dimensions must be fully
  qualified: `"D1@Sketch2@Part1.SLDPRT"`. In an assembly, a component's feature is
  `"Plane1@Part1-1@Assem1"`. **Faces and edges can't be selected by name; they need X, Y, Z**,
  which is why we use the geometry search above.
- Marks matter: loft profiles = 1, sweep path = 4, mate entities = 1, shell faces = 1.

**Exact sizes** [PROVEN]. `GetBodyBox` and `GetPartBox` are loose on curved bodies (a helix's box
is too big), so they can't check an OD to 0.001 mm. Use the extreme point along a direction:

```python
def extent(bodies, d):
    """Furthest point (mm) of the bodies along unit direction d."""
    best = -1e9
    for b in bodies:
        ok, x, y, z = cast(b, "IBody2").GetExtremePoint(*d)      # late-bound: same tuple
        best = max(best, (x * d[0] + y * d[1] + z * d[2]) / MM)
    return best
od = extent(bodies, (1, 0, 0)) + extent(bodies, (-1, 0, 0))
```

For an inner diameter, sample `face.GetClosestPointOn(x, y, z)` from points on the axis and take
the smallest radius.

### 9.1 Measuring an existing part with rays [PROVEN]

Before adding features to a part someone else modelled, measure it. Never trust one nominal plane:
on drafted or curved walls the skin moves.

```python
bodies = list(part.GetBodies2(SWC.swSolidBody, True))
n = doc.RayIntersections(dispatch_array(bodies),
                         darr([c for p in starts for c in p]),    # metres, flat x,y,z list
                         darr([c for d in dirs for c in d]),      # unit vectors
                         SWC.swRayPtsOptsNORMALS | SWC.swRayPtsOptsENTRY_EXIT, 0.0, 0.0)
hits = doc.GetRayIntersectionsPoints()     # 9 doubles per hit: body, ray, type, x, y, z, nx, ny, nz
```

- Batch up to ~4,000 rays per call. This gives the exact outer surface, its normal, the wall
  thickness and any internal features. With `ENTRY_EXIT`, a ray through a 5 mm block gave two hits:
  type **17 at entry** and **33 at exit**. A ray that misses gives no hit. (`swlib.ray_hits` returns
  `(ray, x, y, z, nx, ny, nz, type)` in mm.)
- **Scan behind each planned footprint**: start rays just inside the inner face, going inward. That
  found a 22 mm deep internal rail that was invisible from outside.
- **In an assembly**, `IComponent2.GetBodies3` returns bodies in the **part's** frame. Use
  `body.Copy()` then `ApplyTransform(comp.Transform2)` before casting rays in assembly space. To go the
  other way: `p_asm = p_part @ R + t` (row vectors), `p_part = (p_asm − t) @ R.T`, directions `d @ R.T`
  (one housing was x/z flipped with a 0.09° tilt).
- **Check every measuring ray**: assert that the hit distance is about what you expected and the wall
  is about nominal, and abort otherwise. One ray started 20 mm *inside* the part, so the dry run cut
  away a large chunk.
- **Project surface normals into the working plane** before using them for ray starts or directions.
  A normal with a small z component drifted the start, and a feature landed 1.9 mm low, into a
  speaker.
- **Clearance to other components**: compare the new part with the untouched original using **paired
  rays** (so only *added* material counts), and test against cylinders for fasteners. Component
  bounding boxes over-report (a screw's box includes its head).
- Added inner material must stay **≥ 2 mm under the measured skin** over its whole footprint. An
  inner pad offset from one nominal plane came within 0.6 mm of a drafted wall and left sliver faces
  outside.

---

## 10. Assemblies

### 10.1 The project pattern (reference build: 59 components, 174 mates, 0 errors) [PROVEN]

1. **Parameters file** (`<project>_params.py`): every dimension, part number, path and colour.
2. **Skeleton part** `<PN>_Skeleton.SLDPRT`: only named reference planes (`PL_Z_SPLIT`,
   `PL_Z_BOARD_BOT`, ...) and axes (`AX_N1`, `AX_H1`, ...) at the design coordinates.
3. **Part scripts** (one per part family): each part is modelled **in the product frame**, so
   its own Front/Top/Right planes coincide with the skeleton's when the part sits at the origin.
4. **Purchased and standard parts**: each gets its own mating datums (`AX` axis, `PL_SEAT`,
   `PL_END0` planes) added after import.
5. **Assembly script**: insert the skeleton first (it comes in fixed), then each component at its
   exact transform, then mates to the skeleton from a table of recipes.
6. **Check script**: constraint status, placement vs plan, What's Wrong, interference.
7. **build_all.py**: runs the steps as subprocesses (`python -P`), closes only the project's
   documents between steps, and accepts step names (`front back asm check`).

### 10.2 Inserting components at an exact pose [PROVEN]

```python
res = app.OpenDoc6(path, SWC.swDocPART, SWC.swOpenDocOptions_Silent, "", 0, 0)   # load it first
app.ActivateDoc3(asm_title, False, 0, 0)                                           # back to the assembly
comp = cast(asm.AddComponent5(path, 0, "", False, "", x*MM, y*MM, z*MM), "IComponent2")
# rows = images of the component's x, y, z axes; t = origin; 16 values: 9 rot, 3 trans, scale, 3 unused
arr = [*rows[0], *rows[1], *rows[2], x*MM, y*MM, z*MM, 1.0, 0, 0, 0]
comp.Transform2 = cast(mu.CreateTransform(darr(arr)), "IMathTransform")
```

Placing a component exactly first and then mating means the mates only lock it, they don't move
it. Afterwards, check that the worst deviation of `Transform2.ArrayData[9:12]` from the plan is
0.000 mm.

- **Open files without stealing the user's documents**: `app.GetOpenDocumentByName(path)` first.
  If the user already has the file open, reuse it and never close it.
- **Insert a component in a chosen configuration** [PROVEN]: pass the configuration name as
  ExistingConfigName, then check it, because it can come in with the active one:

```python
comp = cast(asm.AddComponent5(path, SWC.swAddComponentConfigOptions_CurrentSelectedConfig, "", False,
                              "Piece A", x*MM, y*MM, z*MM), "IComponent2")
if comp.ReferencedConfiguration != "Piece A":
    comp.ReferencedConfiguration = "Piece A"           # then read it back again
```

- A `Transform2` you set is **not kept** if mates solve the component somewhere else: the next
  rebuild snaps it back to the mate solution. Use `Transform2` only to place before mating, never to
  test motion (10.8).

### 10.3 Mates with AddMate5 (works on 2025) [PROVEN]

```python
# AddMate5 argument positions (1-based): 1 type, 2 align, 3 flip, 4 Distance, 5 DistanceAbsUpperLimit,
# 6 DistanceAbsLowerLimit, 7-8 gear ratio, 9 Angle, 10 AngleAbsUpperLimit, 11 AngleAbsLowerLimit (RADIANS),
# 12 ForPositioningOnly, 13 LockRotation, 14 WidthMateOption, 15 ErrorStatus (ByRef)
def mate(e1, e2, name, mtype=SWC.swMateCOINCIDENT, align=SWC.swMateAlignCLOSEST, flip=False,
         dist=0.0, lo=None, hi=None, angle=0.0, ang_lo=None, ang_hi=None):
    """Lengths mm, angles deg. A limit mate gets lo/hi (distance) or ang_lo/ang_hi (angle)."""
    doc.ClearSelection2(True)
    if not (e1.Select2(False, 1) and e2.Select2(True, 1)): raise RuntimeError(f"{name}: selection")
    r = math.radians
    res = asm.AddMate5(mtype, align, flip, dist*MM, (dist if hi is None else hi)*MM,
                       (dist if lo is None else lo)*MM, 1, 1,
                       r(angle), r(angle if ang_hi is None else ang_hi), r(angle if ang_lo is None else ang_lo),
                       False, False, 0, 0)          # late-bound: last arg VARIANT(VT_BYREF|VT_I4, 0)
    m, err = res if isinstance(res, tuple) else (res, None)
    if m is None: raise RuntimeError(f"{name} failed, swAddMateError {err}")   # 1 = NoError
    f = mates()[-1]; f.Name = name                  # the new mate = last sub-feature of the MateGroup
    doc.ClearSelection2(True); return f

def newest(doc, type_name):
    """Last feature of a type in the tree. In an assembly the newest feature is NOT
    FeatureByPositionReverse(0): the MateGroup comes after the reference geometry."""
    out, f = None, doc.FirstFeature()
    while f is not None:
        f = cast(f, "IFeature")
        if f.GetTypeName2() == type_name: out = f
        f = f.GetNextFeature()
    return out

def mates():
    sub, out = cast(newest(doc, "MateGroup"), "IFeature").GetFirstSubFeature(), []
    while sub is not None:
        sub = cast(sub, "IFeature"); out.append(sub); sub = sub.GetNextSubFeature()
    return out

def mate_errors():
    """Names of mates with an error after a rebuild (0 = OK)."""
    return [m.Name for m in mates() if m.GetErrorCode2()[0]]   # late-bound: GetErrorCode2(VARIANT(VT_BYREF|VT_BOOL, False))

f_comp = cast(comp.FeatureByName("Front Plane"), "IFeature")   # entities = component features
f_skel = cast(skel.FeatureByName("PL_Z_SPLIT"), "IFeature")
```

Mate recipes we use:

| Component | Mates (all to the skeleton) |
|---|---|
| Part modelled in the product frame | Front↔Front (or the named target plane), Top↔Top, Right↔Right, all coincident |
| Sub-assembly with its own frame | Front↔`PL_Z_xxx`, Right↔`PL_X_xxx`, Top↔`PL_Y_xxx` |
| Screw / insert / spacer / pin | `AX`↔skeleton axis (coincident), `PL_SEAT`↔seat plane (coincident), Top Plane ∥ a skeleton plane (**parallel, rotation lock**) |
| Spring end / range of motion | DISTANCE mate with limits: `dist`, `hi = free length`, `lo = solid length` |
| Rotation with end stops (torsion spring leg, hinge) | ANGLE mate with limits, measured against a **90° helper datum** (below) |

**Limit ANGLE mates: never let the range include 0° or 180°** [PROVEN]. An angle mate can't tell
which side it is on at 0° or 180°. A leg range of 0°..52°, measured between two planes that are
parallel when free, either froze the part or acted the wrong way, with `flip=False` and with
`flip=True`. The fix:
1. Add a helper datum at **90°** to one of the references (here, "Leg 2 Normal Plane").
2. The mate is then `angle = 90`, `ang_lo = 90`, `ang_hi = 90 + θmax` (142° for 52°). That stops it
   at both ends: opening and over-winding.
3. Test it with the drag test (10.8): drive to the middle, the limit, past the limit and the wrong
   way. Check it is blocked at both ends. If it acts the wrong way, delete it and retry with
   `flip` toggled.

- The first component inserted is **fixed**. If you then mate it to the assembly planes, unfix it
  first (`Select4` → `UnfixComponent()`). Otherwise those mates are redundant or over-defining.
- If a mate comes out flipped, delete it and recreate it with the other alignment
  (ALIGNED/ANTI_ALIGNED). Don't stack another mate on top.
- **Mates check themselves in the same run.** A travel mate solved to −40 mm instead of +40 mm, and
  fixing `flip` by hand cost an assembly rerun. Give every mate whose direction matters a check:
  after the mate, rebuild and read the component's position (`Transform2.ArrayData[9:12]`) or angle.
  If it is wrong, delete the mate and add it again with `flip` toggled (then with the other alignment).
  Only if both fail, stop with an error. `swlib.Asm.mate(..., verify=lambda: ...)` does this.
- Group components into folders: select them, then `fm.InsertFeatureTreeFolder2(swFeatureTreeFolder_Containing)`
  and rename the new `FtrFolder`.

### 10.4 CreateMate: the officially recommended API [DOC + PROVEN for WIDTH]

The 2025 help marks `AddMate5` **obsolete, superseded by `IAssemblyDoc::CreateMate`**. It still
works, but use CreateMate for new code, and you must use it for width, symmetric and mechanical
mates.

```python
md = asm.CreateMateData(SWC.swMateCOINCIDENT)          # returns the type-specific *MateFeatureData
md.EntitiesToMate = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [e1, e2])
md.MateAlignment = SWC.swMateAlignALIGNED
feat = asm.CreateMate(md)                              # returns IFeature (None = failed)

# Width mate (centred): proven in Spring_Assembly_build.py
md = asm.CreateMateData(SWC.swMateWIDTH)
md.WidthSelection = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [w1, w2])
md.TabSelection   = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [t1, t2])
md.ConstraintType = SWC.swMateWidth_Centered
asm.CreateMate(md)
```

Mate data interfaces: IAngle-, ICamFollower-, ICoincident-, IConcentric-, IDistance-, IGear-,
IHinge-, ILinearCoupler-, ILock-, IParallel-, IPerpendicular-, IProfileCenter-, IRackPinion-,
IScrew-, ISlot-, ISymmetric-, ITangent-, IUniversalJoint-, IWidthMateFeatureData. To edit a mate:
`feat.GetDefinition()`, change the data, then `feat.ModifyDefinition(...)`. A **symmetric mate was
rejected** in our tests. Use a centred **width** mate for "keep evenly spaced".

### 10.5 Fasteners from Toolbox [PROVEN]

Toolbox parts are configured masters. If you insert one directly, SOLIDWORKS **silently redirects
it to the master's active size** (our M2.5×16 screws loaded as M2.5×10). Instead:
1. Open the master read-only (`swOpenDocOptions_ReadOnly`), then `SaveAs3` it to a new file per
   size (`IS2269_SHCS_M2.5x16_12.9_Black.SLDPRT`).
2. Set `ext.ToolboxPartType = swNotAToolboxPart` and delete the other configurations.
3. Set the size dimensions: `cast(doc.Parameter("Length@BodySke"), "IDimension").SetSystemValue3(L*MM, swSetValue_InAllConfigurations, None)`
   (also `Thread_length@BodySke`), then rebuild and rename the configuration.
4. Add `AX` (axis) and `PL_SEAT` (head underside plane), plus custom properties: Standard,
   Supplier, SKU.

### 10.6 Vendor STEP models [PROVEN]

```python
prev = app.GetUserPreferenceToggle(SWC.swMultiCAD_Enable3DInterconnect)
app.SetUserPreferenceToggle(SWC.swMultiCAD_Enable3DInterconnect, False)    # native features, not a link
try:
    data = app.GetImportFileData(step); res = app.LoadFile4(step, "r", data, 0)
finally:
    app.SetUserPreferenceToggle(SWC.swMultiCAD_Enable3DInterconnect, prev)
doc = cast(res[0] if isinstance(res, tuple) else res, "IModelDoc2")
ext.SaveAs3(out_sldasm, 0, SWC.swSaveAsOptions_Silent | SWC.swSaveAsOptions_SaveReferenced, None, None, 0, 0)
```

- Use the **real** vendor models (GrabCAD, manufacturer STEP, KiCad PCBA STEP). Record every
  source URL in `SOURCES.md`, and keep the raw downloads untouched in an `Online CAD/` folder. The
  built-in browser pane can't save files, so the user downloads them.
- **Keep assemblies light.** A vendor board had 1,082 instances. We pre-processed the STEP with
  OCP (pythonocc / cadquery-ocp): dropped parts smaller than 2.5 mm and replaced parts with more than
  500 faces by their exact bounding-box envelope (`_ENV`). That took it to 18 parts / 3.7 MB.
- Don't re-download parts that are already in the PCB's KiCad STEP export.

### 10.7 Flexible sub-assemblies (springs, mechanisms) [PROVEN]

For a part that must deform in the parent assembly, build it as a sub-assembly of rigid pieces
held by mates. Use this when equations aren't wanted, since an equation-driven part length needs
a rebuild on every move. Two patterns we built:
- **Compression spring**: coil turns kept evenly spaced by width mates between End A and End B,
  with a limit distance mate between the ends.
- **Torsion spring**: two configurations of one part (7.11). Piece A is coil + leg 1. Piece B is
  leg 2, coincident with the spring axis and the coil end plane, which leaves one rotation. A limit
  angle mate through a 90° datum (10.3) sets the range. Winding up is the direction the wire is
  wound: for a left-hand spring advancing +Y, negative about +Y.

**Make Subassembly Flexible through the API** [PROVEN]:

```python
com_put(comp, "Solving", SWC.swComponentFlexibleSolving)   # makepy has no setter; late-bound: comp.Solving = 1
asm_doc.ForceRebuild3(False)
if comp.Solving != SWC.swComponentFlexibleSolving: raise RuntimeError("not flexible")
```

**Rebuild order**: a parent assembly holds its sub-assemblies and parts open. Before rebuilding a
child file, close the parents first, in order (test rig, then flex sub-assembly, then part).
Otherwise the child is stale or locked.

### 10.8 Drag test: prove the motion and the limits [PROVEN]

Setting `Transform2` and rebuilding doesn't simulate a drag, because the component snaps back to
the mate solution. To "drag" a component to angle X:
1. Add a temporary assembly axis (two assembly planes, `InsertAxis2`) and a temporary plane at X
   (`InsertRefPlane(Angle, rad, Coincident, 0, 0, 0)` with Right Plane and that axis selected).
   Read its normal back. If it is reversed by 180°, swap the alignment of the hold mate (7.4).
2. Add a **parallel "hold" mate** from the component's plane to it, `ForceRebuild3`, then read
   the component's `Transform2` and `mate_errors()`.
3. Delete the hold mate and the temporary plane. At the end, delete the axis.

Run it for free, half, the limit, past the limit (it must be blocked: an error, or the hold mate is
rejected) and the wrong way (blocked). Then drive it back to free and check the angle is 0. Print
a table of target vs actual angle for every row in the report.

---

## 11. Speed and stability

- Turn these off during a build and **always turn them back on** (in `finally`):
  `cast(doc.ActiveView, "IModelView").EnableGraphicsUpdate = False`, `fm.EnableFeatureTree = False`,
  `sm.AddToDB = True`, `sm.DisplayWhenAdded = False`. From an external (out-of-process) Python
  script, `app.CommandInProgress = True` also reduces screen updates [DOC].
- **A call that hangs means a dialog is open.** We have seen: the "SOLIDWORKS CAM Warning" on
  opening a file, "Save changes?" on close, and the Modify-dimension pop-up. If a call stalls for
  more than ~30 s, look at the SOLIDWORKS window (screenshot) before retrying. Run long builds in
  the background.
- **Leave the user's documents alone** [PROVEN]. Record the titles of the documents your script
  creates, and close only those on exit, including when it crashes:

```python
_created = []                                   # untitled docs this script made
atexit.register(lambda: [app.CloseDoc(t) for t in reversed(_created)])
doc = cast(app.NewDocument(template, 0, 0, 0), "IModelDoc2"); _created.append(doc.GetTitle())
# after a successful SaveAs3 the file is the user's: stop tracking it (leave it open)
_created.remove(title)
```

  Open existing files with `GetOpenDocumentByName` first, so one the user has open is reused and
  never closed. After rendering, reactivate the document that was active before.
- **Progress you can see, runs you can stop** [PROVEN]:
  - Run build scripts with `python -u` (unbuffered) and print a timestamp per feature. Never pipe
    through `| tail`: that hid 20 minutes of progress.
  - Anything over ~2 min runs in the background (`run_in_background`); tool calls time out at 600 s.
    Wait with a background until-loop that checks for a condition, not with `sleep`.
  - To see whether SOLIDWORKS is busy: `Get-Process SLDWORKS` (CPU, `MainWindowTitle`, `Responding`).
  - To stop a run, kill the real interpreter. The WindowsApps `python` launcher starts a child
    `python3.11.exe`, so match that process *and* the script name in its command line.
- **Build in groups** that can be rebuilt alone (one function per feature group, plus its own remove
  step), so a change to one group doesn't need the whole part.
- Expected timings: a part takes 5–60 s. A full product rebuild (10 parts, fasteners, imports,
  59-component assembly, interference check) takes several minutes.

### 11.1 Troubleshooting playbook

Find the row for the symptom and run its first check before changing any code.

**Connection and crashes**

| Symptom | First check | Likely cause | Fix |
|---|---|---|---|
| `com_error -2147023179` *The interface is unknown* on the first call | `Get-Process SLDWORKS`; is the splash screen up? | SOLIDWORKS is still starting | wait and retry; `swlib.SwApp()` retries up to ~4 min |
| `com_error -2147023174` *The RPC server is unavailable* on any call | `Get-Process SLDWORKS`; SOLIDWORKS's crash dialog names the fault module | the process the script was attached to crashed or exited | **stop every run**, no retry can work. The user restarts SOLIDWORKS and accepts recovery; then check the next row before attaching [PROVEN 2026-10-31] |
| A call hangs for more than ~30 s | screenshot of the SOLIDWORKS window | a modal dialog (section 11) | close the dialog, or ask the user to |
| After a crash or restart, scripts act on the wrong window | `Get-Process SLDWORKS \| Select Id, Responding, MainWindowTitle` | more than one SLDWORKS (one may have no window title); `GetActiveObject` takes whichever registered first | ask the user which to keep; never kill one yourself |
| SOLIDWORKS crashes when a second script process reads what the first wrote | did the second process call `DefineAttribute` for a name the session already has? | not proven yet | verify in the writing process (8); cross-process checks only on request, work saved |

**Return codes**

| Symptom | First check | Likely cause | Fix |
|---|---|---|---|
| An error number that isn't in the enum (e.g. `SaveAs3` → `(False, 2049, 0)`) | split it into bits (2049 = 2048 + 1) and list the swconst names with those values: `{k: v for k, v in SWC.__dict__.items() if isinstance(v, int) and v in (1, 2048) and ("Save" in k or "Path" in k)}` | error returns are **bit masks**: 1 `swGenericSaveError` + 2048 `swFileSaveAsNameExceedsMaxPathLength` | here: a shorter output path [PROVEN] |
| Save error 1 on the second run to the same path | is the earlier result still open? | the old title no longer matches after `SaveAs3` (8) | close the open copy by path before saving |

**A documented call returns None or does nothing**

| Symptom | First check | Likely cause | Fix |
|---|---|---|---|
| A method returns None or has no effect from a script | a Test Run with one OK line per step (create, set, save, reopen, read) for each candidate method | the method needs an add-in context (`IGet3rdPartyStorage`), the other binding mode (3.0), or an active document/selection | use the method the Test Run proved; write the result into the skill |

---

## 12. Verification checklist (run after every build)

| Check | How |
|---|---|
| Rebuild errors | `doc.ForceRebuild3(False)`, then `ext.GetWhatsWrongCount() == 0` |
| Sketches fully defined | `sketch_feat.GetSpecificFeature2().GetConstrainedStatus() == 3` |
| Each feature did something | volume before vs after (`ext.CreateMassProperty().Volume`); body count stays 1 |
| Geometry matches the plan | exact extents with `GetExtremePoint` in 6 directions (9); `GetBodyBox` / `GetPartBox` only for flat-faced parts; volume vs hand calculation (e.g. frustum formula, or wire length × πd²/4 for a spring). Our FOV and shroud parts matched to 0.1 mm³ |
| Parametric | change variables (`set_global`, 7.8), rebuild, What's Wrong = 0, restore |
| Every configuration, just before saving | for each: activate, body count, volume, What's Wrong = 0, suppression table (7.11). Then activate the main configuration |
| Mates | `mate_errors()` empty after the rebuild (10.3) |
| Motion and limits | drag test table (10.8): moves where allowed, blocked past the limits |
| Assembly constraints | every `comp.GetConstrainedStatus() == 3` (the skeleton is fixed) |
| Placement | `Transform2.ArrayData[9:12]` vs the plan, worst deviation ≈ 0 |
| Interference | `idm = asm.InterferenceDetectionManager`; set `TreatCoincidenceAsInterference=False`, `TreatSubAssembliesAsComponents=True`, `IncludeMultibodyPartInterferences=False`, `IgnoreHiddenBodies=True`, `UseTransform=True`; `idm.GetInterferences()`; classify each against an **INTENDED** list (thread engagement, insert press-fit, crush ribs, compressed foam) and report the rest as clashes; `idm.Done()`. **Run it twice**: with `TreatCoincidenceAsInterference=False` (must be zero clashes), then with `True` to list the touching contacts. Each contact must be an intended one (a leg resting on a peg). With `TreatSubAssembliesAsComponents=True` a sub-assembly counts as one component, so overlaps *inside* it are not reported (two overlapping pieces in a flex sub-assembly gave zero results from the parent). Check each sub-assembly on its own as well |
| Exports exist | list the output folder, report the file sizes |
| Changed an existing part | **skin diff** against the untouched original, every build: rays from outside over the changed area + 6 mm. "Raised" points are added material poking out (slivers); "cut outside" points are stray cuts. Both must be 0. Volume returns exactly to the original when a group is removed |
| Openings that must be light-tight or blocked | sight-line test on the built B-rep: thousands of random rays at each opening (±85° across, ±60° along). A leak is a ray that crosses the inner plane **inside** the opening's footprint; rays that miss the part are not leaks. Open-channel test: rays along each slot leg travel its full length without a hit. Prove it in 2D first (a section trace), then on the B-rep |
| Drawings | dry run (every edge found), PDF collision check and the GD&T checklist (8.3); re-run only the failing sheet |

The interference check found real design errors in our builds: a switch cap socket that didn't
clear the plunger, USB-C openings sized to the bare connector instead of the receptacle front, a
port centre 0.3 mm off, Toolbox size redirection, and clamp jaws hitting phone buttons. Always run
it.

---

## 13. VBA macros (.swp)

`.swp` is a binary file. A script can't write one as text. Options:
1. **Preferred:** drive SOLIDWORKS directly from Python (everything above). It's re-runnable,
   versionable and checkable.
2. Write VBA source to a `.bas` file. The user creates a macro (Tools > Macro > New), then
   File > Import in the VBA editor, and runs it (Tools > Macro > Run).
3. Run an existing macro from Python:
   `app.RunMacro2(r"C:\path\macro.swp", "Module1", "main", SWC.swRunMacroUnloadAfterRun, 0)`. [DOC]

Official workflow: plan, then record (Tools > Macro > Record; record twice to keep out view
changes), then edit (delete extra lines, **declare and early-bind variables**), then run, then
debug, then optionally assign to a toolbar button. VSTA macros (VB.NET/C#) build to `.dll` and run
like `.swp`. Macros need references to the SOLIDWORKS 2025 Type Library, Constant type library
and Commands type library.

```vba
Option Explicit
Dim swApp As SldWorks.SldWorks

Sub main()
    Set swApp = Application.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.NewDocument(swApp.GetUserPreferenceStringValue(swDefaultTemplatePart), 0, 0, 0)
    swModel.Extension.SetUserPreferenceInteger swUnitSystem, 0, swUnitSystem_MMGS
    Dim ok As Boolean
    ok = swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
    swModel.SketchManager.InsertSketch True
    swModel.SketchManager.CreateCenterRectangle 0, 0, 0, 0.0125, 0.012, 0     ' metres
    swModel.SketchManager.InsertSketch True
    Dim f As SldWorks.Feature
    Set f = swModel.FeatureManager.FeatureExtrusion3(True, False, False, swEndCondBlind, 0, 0.0025, 0, _
        False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)
    If f Is Nothing Then MsgBox "Extrude failed": Exit Sub
    f.Name = "Flange"
End Sub
```

`RecordMacro` / `EditMacro` aren't available through IDispatch (original skill, SW 2023). Create
a blank `.swp` by hand if you need one.

---

## 14. Design rules learned on our products (3D-printed enclosures, springs)

- **Part numbers and folders**: `<PRODUCT>.<ASM>.<PART>.<REV>`, e.g. `ENC.01.01.00` front shell,
  `ENC.01.02.00` back shell. Project layout: `01_Printed_Parts/ 02_Electronics/ 03_Hardware/
  04_<Mount>/ Scripts/ Outputs/ References/` plus `<PN>_Design_Plan.md` (decisions and corrections)
  and `<PN>_Project_Guide.md` (how to open, change and rebuild).
- **Walls** 2.0 mm PETG. Fillet the outer perimeter (R2) before shelling. Use a tongue-and-rebate
  split line (0.2 gap).
- **Heat-set inserts**: use one size everywhere (M2.5×4, OD 3.88). Hole Ø3.3–3.6 depends on the
  printer, so **print a test coupon first** and put the result into the parameters.
- **Clearances**: 0.2 mm crush ribs for press-held parts. Size connector openings to the real
  receptacle front plus about 0.6 mm (USB-C 9.58×4.16 → 10.2×4.8). Plunger sockets get +0.4 mm.
- **Fasteners**: real standard parts (Toolbox IS 2269 / ISO 4762) and supplier SKUs
  in the custom properties.
- **Optics**: model the camera field of view as a rectangular pyramid frustum from the sensor size
  and focal length (`half = atan(sensor/2/f) + margin`), not a cone. It makes the window size and
  mounting orientation visible.
- **Open items**: always list what must be measured with calipers before printing.

### 14.1 Torsion springs (from the torsion-spring build) [PROVEN]

- **Body turns**: close-wound, the body length is L = (N + 1)·d, so **N = L/d − 1** (8 turns for
  L = 9d, not 9).
- **Leg angle**: with both legs leaving from the same angular point at opposite ends, 180° apart,
  N must be a whole number. Otherwise the free angle changes with N.
- **Rate**: k = E·d⁴ / (64·D·N_eff) in N·mm/rad, with **N_eff = N + (H₁ + H₂)/(3π·D)** for straight
  legs of length H. Shigley's version with friction, per turn: k′ = d⁴E / (10.8·D·N_eff).
- **Maximum angle** (static, 302 stainless): Sut = 1867 / d^0.146 MPa, σ_allow = 0.61·Sut.
  M = σ_allow·π·d³ / (32·Ki), with C = D/d and Ki = (4C² − C − 1) / (4C(C − 1)). θmax = M / k.
- **Wound up**: the mean diameter shrinks to D·N / (N + θ/360°), so ID = that − d. The body
  grows by about d·θ/360°. Check the pin clearance **at the limit**, not only when free.
- **Model**: a helix (7.6, check the hand) with a circular-profile sweep. Leg paths start 0.3 mm
  back inside the wire (overlap), and leg 2 is a separate body joined at the end (7.11). Use
  global variables `T, ID, L, H, N = int(L/T − 1 + 0.5)` and put the rate, max angle and torque
  into the custom properties.
- **Flexible sub-assembly and test rig**: 10.7, 10.8.

### 14.2 Precision parts and instruments (from the caliper build) [PROVEN]

- Fit and clearance table first, standard values computed in the parameters file (rule 1.17).
  Example: ISO 13385-1 caliper accuracy rounds **up** to 10 µm, so it is ±0.03 mm for every length
  above 0 on a 0.02 mm vernier. The hand-typed "±0.02 for 0–100 mm" was wrong.
- Engraved scales: one line plus a pattern, engraved last, behind a `NO_MARKING` configuration (7.12).
- Materials missing from `SOLIDWORKS Materials.sldmat`: AISI 420 → *Chrome Stainless Steel*,
  AISI 301 / 303 → *AISI 304* (AISI 302 → *AISI 304*, 7.10). State the real grade in the drawing's
  material field and in the custom properties. To list the library, read the `.sldmat` file as
  **UTF-16** XML.

### 14.3 Vents, patterns and print orientation (FDM, 0.4 mm nozzle) [PROVEN]

- **Light-tight vent slot (chevron)**: slot width W at the face, legs at angle θ, depth T. The sideways
  jog is J = (T/2)·tan θ. No straight line gets through if J > W. Used: W 1.52, θ 38°, internal gap 1.2,
  rib 1.48, pitch 3.4, T 5.6, giving J 2.19. It stayed light-tight with the slots printed 0.4 mm
  oversize. Extruded along the print Z, it prints with no supports; slot ends are R = W/2 arches.
- A cut's sideways jog must point **away** from neighbouring openings (one broke into a USB window and
  leaked light).
- A backing pad or reinforcement is the **rectangular envelope of the cut plus a margin**, never the
  cosmetic outline (a pill-shaped pad left the end-slot corners unbacked).
- **Flush vents on curved skins**: each slot first goes straight in (depth = the skin's height
  variation over that slot + 0.15 mm), then the chevron starts. The openings stay clean. A flat floor
  and per-slot depths both still leaked, so start with the straight entry.
- **Patterns interrupted by an existing feature**: lay the pattern out **symmetric about that
  feature**, measured from its real edges. A grid with a band removed gave 2 slots on one side and 3
  on the other, with 8.6 mm and 3.0 mm gaps.
- Graduated slot lengths that follow a pill outline (the pill's chord minus 2 × margin) give a
  designed look.
- Minimums: slot 1.5 mm at the face, internal gap ≥ 1.2 mm (3 line widths), ribs ≥ 1.45 mm (3–4
  perimeters). Light-trap pad ≥ 2 mm under the outer skin, ≥ 0.4 mm clear of PCBs, ≥ 1 mm clear of
  displays and screws.
- **Print orientation**: analyse it from the STL with numpy alone. Test the 6 axis orientations for
  height, bed contact area and overhang area (n_z < −0.707). Render with a point-splat z-buffer
  (1.5–2 M area-weighted samples), with overhangs in red, **from above and from below**. Always look at
  the underside before choosing; one choice was reversed after seeing it. Housings: open (joint) side
  down, cosmetic face up as a top surface (ironing), tree supports on the build plate only, 0.16 mm
  layers, 4 walls, ASA on an enclosed printer rather than PLA.

### 14.4 Domain example kept from the original skill: shield machine cutterhead (盾构机刀盘)

Typical features: **face plate** (large disc Ø6–15 m, 400–800 mm thick, extruded), **spokes/arms**
(4–8 radial, rectangular section), **disc-cutter mounts** (circular pattern), **scrapers/rippers**
(on the edge), **centre cutter**, **muck/discharge openings** (through cuts), **mixing bars** (on the
back face), **mounting flange** (central connection). Units: a 6 m diameter is `6.0` in the API.
Build it as: face-plate extrude, then one spoke and circular pattern, then one cutter seat and
circular pattern (or a pattern driven by a table/sketch for the spiral layout), then through-cut
openings, then the flange revolve.

---

## 15. Working with the user

- Make a short plan and confirm assumptions (sizes read off photos, unknown coil counts, etc.)
  before building. Mark guesses as guesses in the report.
- **Before modelling**, do the following once and get it right: a research summary, the parameters
  file, the fit and clearance table, and the computed standard values (rule 1.17).
- Confirm the style with a quick render before any CAD: recessed, flush or chamfered, the layout,
  and symmetry about existing features (rule 1.16). Recessed vent pockets designed without asking,
  when the user wanted flush, cost a full redesign.
- When a drawing letter or callout can mean two things ("B": a bent tip or a straight leg?), don't
  settle it with a text question alone. Build **only the part**, send a render with
  `SendUserFile`, and get a yes before you build the sub-assemblies and rigs (rule 1.16). A misread
  "B" cost one full rebuild of part, flex sub-assembly and test rig.
- Build in stages that the user can see: part → render → sub-assembly → test rig. Each stage is a
  separate script that can be re-run alone (`build_all.py part flex test`).
- Reuse the existing proven code instead of rewriting it (section 16).
- Don't touch the user's open assemblies. Build into new files, and say which files were created.
- Minimise back-and-forth. Do the whole build plus checks in one run, then report:
  files (as links), feature tree, checks with numbers, changes from the spec and why, open items.
- After the user edits something by hand (silkscreen, a sketch, a colour), capture their change
  and keep it on regeneration. Never overwrite it with a redraw.

---

## 16. Proven code to reuse (in this repository)

| What | File |
|---|---|
| **swlib.py**: early-bound API layer (model-coordinate sketching on any plane, verified planes, axes, extrude/cut between coordinates, revolve, shell, fillet/chamfer by geometry, outline text, equations, colours, export, `Asm` with placement + mates) | `examples/swlib.py` |
| swlib 2.1 additions from the torsion-spring build: `NOTHING`, `dispatch_array`, `com_put`, `newest`, plane flip retry, `Part.feature`, `extent`/`extents`, `helix` + `helix_hand`, `sweep_circular`, `combine_add`, `set_global`, configurations (`add_configuration`, `suppress_in`, `suppression`), per-configuration `color`, material read-back, `Asm.add(config=)`, angle-limit `mate`, `mates`/`mate_errors`, `make_flexible`, two-pass `interference`, `SwApp.close_created`, startup wait. All ran early-bound on 2025 SP1.2 (2026-10-30) except the startup wait | `examples/swlib.py` |
| swlib 2.2 additions from the caliper and vent builds: `Asm.mate(verify=)`, save by path (`close_other`), `plane_3pts`, `name_entity` / `entity`, `folder`, `delete_group`, `ray_hits`, `backup`, `log`. All ran early-bound on 2025 SP1.2 (SolidWorks Test Run, 2026-10-31) | `examples/swlib.py` |
| Hidden and signed metadata (an author credit): summary info + custom property + hidden attribute + NTFS stream + HMAC signature + STEP/PDF author (`stamp_credit`, `read_credit`, `stamp_file_ads`, `stamp_step`, `stamp_pdf`; set `SwApp.credit` to stamp on every save). Part, assembly, STEP and PDF checked INTACT (2026-10-31). Check files with `tools/verify_credit.py` (SOLIDWORKS files only with `--sldworks`, see 8) | `examples/swlib.py`, `tools/verify_credit.py` |
| SolidWorks Test Run template (rule 1.14) | `examples/00_test_run_template.py` |
| Parametric plate: sketches, cuts, fillet and chamfer by geometry, global variables, checks, STEP/STL export | `examples/01_parametric_plate.py` |
| Skeleton + parts + pinned assembly with mates, constraint check, interference check | `examples/02_pinned_assembly.py` |
| VBA macro (early bound) to import into a `.swp` | `examples/vba/CreateFlange.bas` |
| Look up any method's exact argument list from your installed type library | `tools/sw_signature.py IFeatureManager.FeatureExtrusion3` |
| Look up enum values from `swconst` | `tools/sw_constants.py swEndCond` |

---

## 17. Sources

- SOLIDWORKS API Help 2025: [CreateMate](https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IAssemblyDoc~CreateMate.html), [ICoincidentMateFeatureData](https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ICoincidentMateFeatureData.html), [Create Standard Mates example (VBA)](https://help.solidworks.com/2025/english/api/sldworksapi/Create_Standard_Mates_Example_VB.htm), [SOLIDWORKS Macros](https://help.solidworks.com/2025/english/api/sldworksapiprogguide/GettingStarted/SolidWorks_Macros.htm), [Run SOLIDWORKS Macro](https://help.solidworks.com/2025/english/api/sldworksapiprogguide/GettingStarted/Run_SolidWorks_Macro.htm)
- [AddMate5 (obsolete → CreateMate)](https://help.solidworks.com/2023/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IAssemblyDoc~AddMate5.html), [FeatureExtrusion3](https://help.solidworks.com/2018/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IFeatureManager~FeatureExtrusion3.html), [SelectByID2](https://help.solidworks.com/2020/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDocExtension~SelectByID2.html), [CommandInProgress](https://help.solidworks.com/2016/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISldWorks~CommandInProgress.html), [EnableFeatureTree](https://help.solidworks.com/2019/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IFeatureManager~EnableFeatureTree.html)
- Installed type libraries (SOLIDWORKS 2025 SP1.2): all signatures in section 6 and constants in section 5.
- Spring formulas: Budynas & Nisbett, *Shigley's Mechanical Engineering Design*, chapter 10 (helical torsion springs).
- Original skill: guoleizhen717/solidworks-api-skill (MIT), translated and corrected here.
