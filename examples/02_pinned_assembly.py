"""Example 2 - skeleton-based assembly: base + cover + two pins, fully mated and checked.

The pattern used for real products (59 components / 174 mates, 0 errors):
1. every part is modelled in the product frame (or with its own named datums: AX, PL_SEAT)
2. a skeleton part holds named datum planes and axes at the design coordinates
3. the assembly inserts the skeleton first (fixed), places every part at its exact pose,
   then mates it to the skeleton - so mates lock parts, they never move them
4. checks: mate errors, every component fully constrained, placement, interference

Run:  python examples/02_pinned_assembly.py
Output: examples/out/pinned_assembly/*.SLDPRT, Pinned_Assembly.SLDASM, Pinned_Assembly.STEP
"""
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import swlib as L                      # noqa: E402
from swlib import SWC, MM, cast        # noqa: E402

# ---------------- parameters (mm) - product frame: origin at the base centre, +z up -------------
BASE_L, BASE_W, BASE_T = 60.0, 40.0, 10.0
COVER_T, CORNER_R = 3.0, 3.0
PIN_X = 20.0                            # pins at (+-PIN_X, 0)
PIN_D, HOLE_D = 4.0, 4.2                # 0.1 mm radial clearance
HEAD_D, HEAD_H = 7.0, 2.0
TOP_Z = BASE_T + COVER_T                # cover top = pin seat

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out", "pinned_assembly")
P = {n: os.path.join(OUT, f"{n}.SLDPRT") for n in ("Skeleton", "Base", "Cover", "Pin")}
ASM = os.path.join(OUT, "Pinned_Assembly.SLDASM")
I3 = ((1, 0, 0), (0, 1, 0), (0, 0, 1))


def plate(sw, name, z0, z1, color):
    """A rounded plate between z0 and z1 with the two pin holes."""
    p = L.Part(sw, name)
    p.graphics(False)
    p.sketch("z", z0, f"SK_{name.upper()}")
    p.rect(-BASE_L / 2, -BASE_W / 2, BASE_L / 2, BASE_W / 2, CORNER_R)
    for sx in (-1, 1):
        p.circle(sx * PIN_X, 0, HOLE_D / 2)
    p.extrude(name.upper(), z1)
    p.color(color)
    p.properties({"Description": f"{name} plate (example 2)"})
    p.save(P[name])
    sw.close(p.doc)


def pin(sw):
    """Pin modelled on its own axis: shank z 0..TOP_Z, head above. Datums AX + PL_SEAT for mating."""
    p = L.Part(sw, "Pin")
    p.graphics(False)
    p.sketch("z", 0, "SK_SHANK"); p.circle(0, 0, PIN_D / 2); p.extrude("SHANK", TOP_Z)
    p.sketch("z", TOP_Z, "SK_HEAD"); p.circle(0, 0, HEAD_D / 2); p.extrude("HEAD", TOP_Z + HEAD_H)
    p.axis("AX", "z", 0.0, 0.0)
    p.named_plane("PL_SEAT", "z", TOP_Z)     # head underside
    p.color((0.78, 0.80, 0.83), shininess=0.6, specular=0.7)
    p.properties({"Description": f"Pin D{PIN_D:g} x {TOP_Z:g} (example 2)"})
    p.save(P["Pin"])
    sw.close(p.doc)


def skeleton(sw):
    p = L.Part(sw, "Skeleton")
    p.graphics(False)
    p.named_plane("PL_Z_TOP", "z", TOP_Z)
    for i, sx in enumerate((-1, 1), 1):
        p.axis(f"AX_P{i}", "z", sx * PIN_X, 0.0)
    p.properties({"Description": "Skeleton: datum planes and axes for all mates"})
    p.save(P["Skeleton"])
    sw.close(p.doc)


def assemble(sw):
    A = L.Asm(sw)
    skel = A.add(P["Skeleton"])                                  # first component = fixed
    comps = []
    for name in ("Base", "Cover"):                               # modelled in the product frame
        c = A.add(P[name], I3, (0, 0, 0))
        for pl in ("Front Plane", "Top Plane", "Right Plane"):
            A.mate(A.feat(c, pl), A.feat(skel, pl), label=f"{name} {pl}")
        comps.append((c, (0, 0, 0)))
    for i, sx in enumerate((-1, 1), 1):                          # fastener recipe: axis + seat + lock
        org = (sx * PIN_X, 0.0, 0.0)
        c = A.add(P["Pin"], I3, org)
        A.mate(A.feat(c, "AX"), A.feat(skel, f"AX_P{i}"), label=f"pin {i} axis")
        A.mate(A.feat(c, "PL_SEAT"), A.feat(skel, "PL_Z_TOP"), label=f"pin {i} seat")
        A.mate(A.feat(c, "Top Plane"), A.feat(skel, "Top Plane"), SWC.swMatePARALLEL, label=f"pin {i} lock")
        comps.append((c, org))
    A.doc.ForceRebuild3(False)
    return A, skel, comps


def verify(A, skel, comps):
    ok = True
    errs = A.ext.GetWhatsWrongCount()
    print(f"  mate / rebuild errors: {errs}"); ok &= errs == 0
    for c, org in comps:
        st = c.GetConstrainedStatus()
        at = [v / MM for v in cast(c.Transform2, "IMathTransform").ArrayData[9:12]]
        dev = sum((a - b) ** 2 for a, b in zip(at, org)) ** 0.5
        print(f"  {c.Name2:12s} constrained status {st} (3 = fully), placement deviation {dev:.4f} mm")
        ok &= st == SWC.swFullyConstrained and dev < 0.01
    idm = cast(A.asm.InterferenceDetectionManager, "IInterferenceDetectionMgr")
    idm.TreatCoincidenceAsInterference = False
    idm.TreatSubAssembliesAsComponents = True
    idm.IncludeMultibodyPartInterferences = False
    idm.IgnoreHiddenBodies = True
    idm.UseTransform = True
    clashes = []
    for it in idm.GetInterferences() or ():
        it = cast(it, "IInterference")
        clashes.append(([cast(c, "IComponent2").Name2 for c in (it.Components or ())], it.Volume / MM ** 3))
    idm.Done()
    for names, vol in clashes:
        print(f"  CLASH {vol:.3f} mm^3: {' <> '.join(names)}")
    print(f"  interferences: {len(clashes)} (expected 0)")
    return ok and not clashes


def main():
    t0 = time.time()
    os.makedirs(OUT, exist_ok=True)
    sw = L.SwApp()
    for path in [ASM, *P.values()]:
        sw.close_path(path)                                      # only our own previous results
    plate(sw, "Base", 0.0, BASE_T, (0.42, 0.46, 0.52))
    plate(sw, "Cover", BASE_T, TOP_Z, (0.93, 0.93, 0.91))
    pin(sw)
    skeleton(sw)
    A, skel, comps = assemble(sw)
    print(f"assembly: {len(comps) + 1} components, {A.nmates} mates")
    ok = verify(A, skel, comps)
    A.doc.ShowNamedView2("*Isometric", 7); A.doc.ViewZoomtofit2()
    A.save(ASM)
    A.ext.SaveAs3(ASM.replace(".SLDASM", ".STEP"), 0,
                  SWC.swSaveAsOptions_Silent | SWC.swSaveAsOptions_Copy, None, None, 0, 0)
    print(f"saved {ASM} (+ STEP) in {time.time() - t0:.0f} s - {'ALL CHECKS PASSED' if ok else 'SOME CHECKS FAILED'}")
    return ok


if __name__ == "__main__":
    sys.exit(0 if main() else 1)
