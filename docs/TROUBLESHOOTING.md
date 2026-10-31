# Troubleshooting

Every entry below was hit in a real build, except where marked *(reported)*.

## Python and COM

| Symptom | Cause | Fix |
|---|---|---|
| `ModuleNotFoundError: No module named 'win32com'` | pywin32 isn't installed, or the interpreter is isolated | `pip install pywin32`. Don't use `python -I` if pywin32 is in your user site. Use `python -P` and add the script folder to `sys.path` |
| `ImportError: No SOLIDWORKS type library found` | SOLIDWORKS isn't installed or not registered | run SOLIDWORKS once as the same Windows user; repair the installation |
| A method "is not callable", or returns a bound method | late binding treats zero-argument methods as properties | use early binding (`swlib.cast`) or the `call(obj, "Name")` helper |
| `'tuple' object is not callable` (`GetSketchSegments()`, `GetConfigurationNames()`, `GetChildren()`) or `'int' object is not callable` (`GetConstrainedStatus()`) | late and early binding mixed in one script: late-bound objects evaluate zero-argument methods as properties | import `swlib` and use early binding everywhere (SKILL.md §3.0). If you must stay late-bound, go through `call(obj, "Name")` for every zero-argument member |
| `Member not found` on `doc.FeatureByName` | `FeatureByName` is on `IPartDoc` / `IAssemblyDoc` / `IComponent2`, not `IModelDoc2` | `cast(doc, "IPartDoc").FeatureByName(...)` (`swlib.Part.feature`) |
| `int() argument must be ... not 'VARIANT'` | a ByRef `VARIANT` passed to an early-bound method | early-bound: pass plain `0` or leave the output out, and read the returned tuple |
| `Type mismatch` when passing `None` (late-bound), e.g. `InsertCombineFeature(SWBODYADD, None, bodies)` | Python `None` isn't a null object | `NOTHING = VARIANT(VT_DISPATCH, None)`, and an object array is `swlib.dispatch_array(objs)` |
| `Type mismatch` on `GetMaterialPropertyName2("", "")` (late-bound) | the database argument is a ByRef string output | late-bound: `VARIANT(VT_BYREF \| VT_BSTR, "")`. Early-bound: `name, db = part.GetMaterialPropertyName2("")` |
| `The server threw an exception` from `IMathUtility.CreatePoint` | late-bound `IMathUtility` | transform points by hand from `ModelToSketchTransform.ArrayData` (SKILL.md §7.3), or cast `IMathUtility` early-bound |
| Can't set `IEquationMgr.Equation(i)`, or delete + `Add2` fails | `Equation` is an indexed property | `swlib.com_put(eq, "Equation", i, '"ID" = 3')` (`Part.set_global`), then `EvaluateAll` |
| `AttributeError` setting `comp.Solving` early-bound | makepy marks `Solving` read-only | `swlib.com_put(comp, "Solving", swComponentFlexibleSolving)` (`Asm.make_flexible`) |
| `The interface is unknown` (-2147023179) on the first call | SOLIDWORKS is still on its splash screen (the script started it) | wait and retry; `swlib.SwApp()` retries for up to 4 minutes |
| `The RPC server is unavailable` (-2147023174) on any call | the SOLIDWORKS process the script was attached to crashed or exited | stop every script; restart SOLIDWORKS and accept document recovery. Then check `Get-Process SLDWORKS` shows only one instance before attaching again |
| SOLIDWORKS crashed while a second script process read files the first had just stamped | not proven; the second process defined the same attribute again in the same session | verify in the writing process; `tools/verify_credit.py` checks SOLIDWORKS files only with `--sldworks`, with your work saved |
| `IGet3rdPartyStorage` returns None | it needs an add-in's save notification; an external script can't use it | use a hidden attribute or an NTFS alternate data stream (skill section 8) |
| `SaveAs3` returns `(False, 2049, 0)` | 2049 = generic error + `swFileSaveAsNameExceedsMaxPathLength` (2048) | use a shorter output path (well under 260 characters) |
| `TypeError` / type mismatch on an array argument | Python lists aren't `SAFEARRAY`s | wrap them: `VARIANT(pythoncom.VT_ARRAY \| pythoncom.VT_R8, values)` (`swlib.darr`) |
| A ByRef output isn't returned | early-bound calls return ByRef outputs as a tuple | `doc, err, warn = app.OpenDoc6(...)`; `mate, err = asm.AddMate5(...)` |
| `SyntaxError: unterminated string literal` after patching a script | a shell heredoc turned `"\n"` in the patch into a real newline | edit scripts with an editor or the Edit tool, never with shell find-and-replace |
## SOLIDWORKS behaviour

| Symptom | Cause | Fix |
|---|---|---|
| The script hangs on a call | a modal dialog is open: SOLIDWORKS CAM warning, *Save changes?*, the *Modify* dimension box | look at the SOLIDWORKS window. Turn off `swInputDimValOnCreate` while building. Close documents you opened with `CloseDoc(title)` |
| An inner rectangle snaps onto the outer one | sketch inference | `SetUserPreferenceToggle(swSketchInference, False)` and `SketchManager.AddToDB = True` (restore afterwards) |
| An offset plane lands on the wrong side | the side isn't guaranteed | read back `IRefPlane.Transform.ArrayData[9:12]`, delete and retry with `swRefPlaneReferenceConstraint_OptionFlip` |
| A boss or cut goes the wrong way | default directions: boss +normal, cut −normal | set `Dir` from the sketch normal; check volume and far-face position; retry reversed |
| A chamfer or fillet feature exists but changed nothing | wrong edges selected | select edges by geometry (`swlib.find_faces`, circle centres); assert the volume dropped |
| A loft fails | outer and inner loops in one sketch | one sketch per loop: loft the outer ones, then loft-cut the inner ones |
| A sweep along a joined helix fails | pitch changes make tiny corners | sweep each helix on its own; they merge into one body |
| A sweep or boss returns `None`, or leaves 2 bodies | it only *touches* the body: face to face, edge to face, or tangent | overlap by 0.2–0.5 mm, or leave a gap of at least 0.01 mm. To debug, re-run with `Merge = False` and print each body's box |
| The helix has the wrong hand | `Clockwised=True` gives a **left-hand** helix (circle on a plane parallel to Top, advancing +Y) | check with `swlib.Part.helix_hand` and stop if it is wrong |
| `OD` check off on a spring or a filleted part | `GetBodyBox` / `GetPartBox` are loose on curved bodies | `swlib.Part.extents()` (`GetExtremePoint` in 6 directions) |
| `SetSuppression2(..., swSpecifyConfiguration, ...)` returns True and suppresses nothing | | activate the configuration, select the features (looked up again by name), `EditSuppress2`, then read back with `IsSuppressed2` (`swlib.Part.suppress_in`) |
| A feature disappears in a configuration where it should stay | its parent is suppressed there | build it as a separate body and join it with a Combine feature at the end of the tree; suppress the join |
| Default has 2 bodies after a temporary suppress | dependents (the Combine) were suppressed too and not unsuppressed | unsuppress the whole list including dependents; check every configuration before saving (`check_configurations`) |
| Other configurations are grey | appearance is per configuration | `swlib.Part.color(rgb, configs="all")` |
| Material is empty after `SetMaterialPropertyName2` | the name isn't in the database (AISI 302 isn't in `SOLIDWORKS Materials.sldmat`); it fails silently | read it back and fall back: `swlib.Part.material("AISI 302", fallbacks=("AISI 304",))` |
| "Axis failed" in an assembly although the axis exists | `FeatureByPositionReverse(0)` returns the Mates folder in an assembly | `swlib.newest(doc, "RefAxis")` |
| A limit angle mate freezes the part or stops the wrong way | the range includes 0° or 180°, where an angle mate can't tell which side it is on | measure against a 90° helper datum: 90°..(90 + θmax) (SKILL.md §10.3) |
| A component set with `Transform2` snaps back on rebuild | mates win over the transform | drag test with a temporary plane and a parallel "hold" mate (SKILL.md §10.8) |
| Component inserted in the wrong configuration | `AddComponent5` got no configuration name | pass it as ExistingConfigName and check `ReferencedConfiguration` (`swlib.Asm.add(config=)`) |
| The interference check finds nothing although parts inside a sub-assembly overlap | `TreatSubAssembliesAsComponents = True` treats the sub-assembly as one component | also run the check inside each sub-assembly |
| A sketch takes many minutes to solve | hundreds of segments in one sketch, all fixed | ≤ 40 segments per sketch; draw one and pattern it; time a 10-segment test run first |
| A drawing or check script crawls through a part | per-face / per-edge COM walks on an engraved part (thousands of faces) | engrave last behind a `NO_MARKING` configuration; name the entities you need (`Part.name_entity`); read only outer loops of non-marking features |
| A mate solves to the wrong side (−40 instead of +40) | flip / alignment | `Asm.mate(..., verify=lambda: ...)` retries flip and alignment in the same run |
| Save fails with error 1 after an earlier run | the earlier file is still open; closing it by its old title did nothing | close by path (`SwApp.close_other`), re-read the title after `SaveAs3` |
| A GD&T symbol prints as raw text (`<IGTOL-PARALLEL>`) | wrong symbol name, or no `ConvertFormat()` | names from `gtol.sym` (`<IGTOL-PARA>`), then `ConvertFormat()` |
| `IDimensionTolerance.SetValues2` returns False on a drawing dimension | | use `SetValues(min, max)` |
| Views off the sheet, notes over the border, texts on texts | annotations placed at hand-picked offsets | plan the sheet zones first, ISO 129 spacing, PDF collision check with `pymupdf` (SKILL.md §8.3) |
| A dry run cut away a large chunk | a measuring ray started inside the part | assert every ray's hit distance and wall thickness before cutting (SKILL.md §9.1) |
| Sliver faces outside after adding inner material | offset from one nominal plane on a drafted / curved wall | stay ≥ 2 mm under the *measured* skin; skin diff against the original after every build |
| A script patched through a shell heredoc is empty or broken | heredoc + non-ASCII text, or `"\n"` turned into a newline | use the Edit tool; `encoding="utf-8"`, `assert old in s` |
| numpy / scipy imports break | a script named like a standard module (`inspect.py`) | rename it |
| A rebuilt sub-assembly is stale or locked | its parent assembly is still open | close parents first: test rig, then sub-assembly, then part |
| Parts look see-through | wrong `MaterialPropertyValues` layout | `[R, G, B, Ambient, Diffuse, Specular, Shininess, Transparency, Emission]`. Index 7 must be 0 |
| Text reads mirrored or rotated on a side face | sketch text follows the sketch frame | use `swlib.label` (right/up vectors) or `swlib.text_outline` (glyph outlines) |
| Toolbox screws come in at the wrong length | Toolbox redirects to the master's active size | save a copy, set `ToolboxPartType = swNotAToolboxPart`, delete the other configurations, set the size dimensions |
| A mate is redundant or over-defining at the first component | the first component is inserted fixed | unfix it before mating it to the assembly planes, or mate others to it instead |
| A symmetric mate is rejected | mate data selection | use a centred **width** mate (`swMateWIDTH`, `swMateWidth_Centered`) |
| A mate flips the part | alignment chosen as CLOSEST | delete it and recreate with ALIGNED / ANTI_ALIGNED; don't stack mates |
| An imported STEP becomes a linked (3D Interconnect) file | 3D Interconnect is on | switch `swMultiCAD_Enable3DInterconnect` off during `LoadFile4`, then restore it |
| The assembly is very slow | vendor models with thousands of faces | replace heavy components with bounding-box envelopes; drop parts smaller than ~2 mm |
| The open document got renamed to `.STEP` after export | `SaveAs` without the Copy option | `swSaveAsOptions_Silent \| swSaveAsOptions_Copy` |

## Still stuck?

1. Print the exact signature: `python tools/sw_signature.py IInterface.Method`.
2. Print the enum: `python tools/sw_constants.py swSomething`.
3. Check `ext.GetWhatsWrongCount()` after `ForceRebuild3(False)`.
4. Open an issue with the SOLIDWORKS version, the Python version, the call and the full traceback.
