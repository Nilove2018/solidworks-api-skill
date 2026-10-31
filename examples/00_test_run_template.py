"""SolidWorks Test Run template: answer ONE question about SOLIDWORKS behaviour in seconds,
before you write or change a full build script (SKILL.md rule 1.14).

A test run is a 20-40 line throwaway script that checks one API behaviour on a dummy part
(here a circle and a helix) in a new, untitled document. It prints the answer and closes only
that document; your open documents are never touched. Write the answer into SKILL.md, then build.

Typical questions: which side an offset plane lands on, the hand of a helix, the direction of a
mate, whether two touching features merge, whether a suppression call applies, and SPEED (time
10 sketch segments or one edge walk before you do 200).

Run:  python -u examples/00_test_run_template.py
"""
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import swlib as L  # noqa: E402


def test_run(sw):
    """Example question: what hand does InsertHelix(Clockwised=True) give, and where does it start?"""
    p = L.Part(sw, "test_run")
    p.sketch("y", 0, "SK_CIRCLE")
    p.circle(0, 0, 5)                               # circle R5 on the Top Plane
    p.end_sketch()
    h = p.helix("SK_CIRCLE", "HELIX", height=10, pitch=2, revolutions=5, left_hand=True)
    hand, start, end = p.helix_hand(h)
    print(f"Clockwised=True -> {hand}; start {[round(v, 3) for v in start]}, end {[round(v, 3) for v in end]}")


if __name__ == "__main__":
    sw = L.SwApp()
    t0 = time.time()
    try:
        test_run(sw)
    finally:
        sw.close_created()                          # only the test run's own untitled document
    print(f"test run took {time.time() - t0:.1f} s")
