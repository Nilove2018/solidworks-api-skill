"""Keep the root SKILL.md identical to the plugin copy.

The skill lives in skills/solidworks-api/SKILL.md (where Claude Code plugins load it from).
SKILL.md at the repository root is a copy for people who browse GitHub or copy one file.

    python tools/sync_skill.py          # copy skills/solidworks-api/SKILL.md -> SKILL.md
    python tools/sync_skill.py --check  # exit 1 if the two differ (use before committing)
"""
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
SRC = ROOT / "skills" / "solidworks-api" / "SKILL.md"
DST = ROOT / "SKILL.md"

if __name__ == "__main__":
    src = SRC.read_bytes()
    if "--check" in sys.argv:
        same = DST.exists() and DST.read_bytes() == src
        print("SKILL.md in sync" if same else "SKILL.md differs - run: python tools/sync_skill.py")
        sys.exit(0 if same else 1)
    DST.write_bytes(src)
    print(f"copied {SRC.relative_to(ROOT)} -> {DST.relative_to(ROOT)} ({len(src)} bytes)")
