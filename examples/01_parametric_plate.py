"""Example 1 - a parametric mounting plate, built and checked through the SOLIDWORKS API.

What it shows
* sketches on the Front Plane and on offset planes, in model millimetres (swlib)
* boss extrudes and through cuts with named features
* a fillet on the outer top edges and chamfers on the hole edges, chosen by geometry
* global variables, custom properties, material and colour
* checks: volume against a hand calculation, bounding box, rebuild errors
* save as .SLDPRT and export .STEP / .STL

Run (SOLIDWORKS open or closed; a new part is created, your open documents are not touched):
    python examples/01_parametric_plate.py
Output: examples/out/Parametric_Plate.SLDPRT / .STEP / .STL
"""
import math
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import swlib as L                      # noqa: E402

# ---------------- parameters (mm) ----------------
PLATE_L, PLATE_W, PLATE_T = 80.0, 50.0, 6.0
CORNER_R = 5.0
HOLE_D, HOLE_INSET = 5.5, 8.0           # 4 corner holes for M5 screws
BOSS_D, BOSS_H, BORE_D = 16.0, 8.0, 8.0  # centre boss with a through bore
EDGE_FILLET, HOLE_CHAMFER = 1.0, 0.5

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")
PART = os.path.join(OUT, "Parametric_Plate.SLDPRT")
HOLES = [(x, y) for x in (HOLE_INSET, PLATE_L - HOLE_INSET) for y in (HOLE_INSET, PLATE_W - HOLE_INSET)]
CX, CY = PLATE_L / 2, PLATE_W / 2


def check(label, got, want, tol):
    ok = abs(got - want) <= tol
    print(f"  {'OK  ' if ok else 'FAIL'} {label}: {got:.3f} (expected {want:.3f})")
    return ok


def hole_edges(p, z, centres):
    """Circular edges on the planar face(s) at height z whose centre is one of `centres`."""
    out = []
    for face in p.find_faces(normal=(0, 0, 1), at=z, axis="z"):
        for e in face.GetEdges() or ():
            e = L.cast(e, "IEdge")
            c = L.cast(e.GetCurve(), "ICurve")
            if not c.IsCircle():
                continue
            cp = c.CircleParams                  # centre x, y, z (m), axis, radius
            if any(abs(cp[0] / L.MM - x) < 0.01 and abs(cp[1] / L.MM - y) < 0.01 for x, y in centres):
                out.append(e)
    return out


def build():
    t0 = time.time()
    sw = L.SwApp()
    sw.close_path(PART)                         # only our own previous result, if open
    L.set_export_prefs(sw)
    p = L.Part(sw, "Parametric_Plate")
    p.graphics(False)
    p.globals({"PLATE_L": PLATE_L, "PLATE_W": PLATE_W, "PLATE_T": PLATE_T, "HOLE_D": HOLE_D})

    # 1. plate with rounded corners, z 0 -> PLATE_T
    p.sketch("z", 0, "SK_PLATE"); p.rect(0, 0, PLATE_L, PLATE_W, CORNER_R); p.extrude("PLATE", PLATE_T)
    # 2. centre boss on top, z PLATE_T -> PLATE_T + BOSS_H
    p.sketch("z", PLATE_T, "SK_BOSS"); p.circle(CX, CY, BOSS_D / 2); p.extrude("BOSS", PLATE_T + BOSS_H)
    # 3. through bore, cut from the boss top down to z 0
    p.sketch("z", PLATE_T + BOSS_H, "SK_BORE"); p.circle(CX, CY, BORE_D / 2)
    p.extrude("BORE", 0, cut=True)
    # 4. corner holes, cut from the plate top down to z 0
    p.sketch("z", PLATE_T, "SK_HOLES")
    for x, y in HOLES:
        p.circle(x, y, HOLE_D / 2)
    p.extrude("CORNER_HOLES", 0, cut=True)

    print("checks before edge treatments:")
    plate = (PLATE_L * PLATE_W - (4 - math.pi) * CORNER_R ** 2) * PLATE_T
    boss = math.pi * (BOSS_D / 2) ** 2 * BOSS_H
    bore = math.pi * (BORE_D / 2) ** 2 * (PLATE_T + BOSS_H)
    holes = 4 * math.pi * (HOLE_D / 2) ** 2 * PLATE_T
    ok = check("volume mm^3", p.volume(), plate + boss - bore - holes, 0.5)
    lo, hi = p.bbox()
    ok &= check("size x", hi[0] - lo[0], PLATE_L, 0.01)
    ok &= check("size y", hi[1] - lo[1], PLATE_W, 0.01)
    ok &= check("height z", hi[2] - lo[2], PLATE_T + BOSS_H, 0.01)

    # 5. edge treatments, edges found by geometry (never by screen picks)
    outer = p.outer_loop_edges(normal=(0, 0, 1), at=PLATE_T, axis="z",
                               rect=(0, 0, PLATE_L, PLATE_W), r=CORNER_R)
    print(f"  outer top edges found: {len(outer)}")
    v = p.volume(); p.fillet("TOP_EDGE_FILLET", outer, EDGE_FILLET)
    ok &= p.volume() < v
    edges = hole_edges(p, PLATE_T, HOLES)
    print(f"  hole edges found: {len(edges)} (expected 4)")
    v = p.volume(); p.chamfer("HOLE_CHAMFERS", edges, HOLE_CHAMFER)
    ok &= p.volume() < v                        # a chamfer must actually remove material

    errors = p.rebuild_check()
    print(f"  rebuild errors (What's Wrong): {errors}")
    ok &= errors == 0

    # 6. finish: properties, material, colour, save, export
    p.properties({"Description": "Parametric mounting plate (example 1)", "Material": "ABS"})
    p.material("ABS")
    p.color((0.55, 0.68, 0.82), shininess=0.4)
    p.save(PART)
    for ext in (".STEP", ".STL"):
        p.export(PART.replace(".SLDPRT", ext))
    print(f"saved {PART} (+ STEP, STL) in {time.time() - t0:.0f} s - {'ALL CHECKS PASSED' if ok else 'SOME CHECKS FAILED'}")
    return ok


if __name__ == "__main__":
    sys.exit(0 if build() else 1)
