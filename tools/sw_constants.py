"""Print SOLIDWORKS enum values (swconst) from the installed type library.

Never hard-code a guessed integer - look it up:

    python tools/sw_constants.py swEndCond           # every constant whose name contains the text
    python tools/sw_constants.py swMateType_e        # every value of one enum
    python tools/sw_constants.py --json > swconst.json

Reads the registered swconst type library directly (no SOLIDWORKS session needed).
Set SW_TLB_MAJOR to pick a version (33 = SOLIDWORKS 2025, 32 = 2024, 31 = 2023).
"""
import json
import os
import sys

import pythoncom

SWCONST_TLB = '{4687F359-55D0-4CD3-B6CF-2EB42C11F989}'


def load_enums():
    forced = os.environ.get("SW_TLB_MAJOR")
    for major in ([int(forced)] if forced else range(45, 19, -1)):
        try:
            tlb = pythoncom.LoadRegTypeLib(SWCONST_TLB, major, 0, 0)
            break
        except pythoncom.com_error:
            continue
    else:
        sys.exit("No SOLIDWORKS swconst type library registered - is SOLIDWORKS installed?")
    enums = {}
    for i in range(tlb.GetTypeInfoCount()):
        ti = tlb.GetTypeInfo(i)
        ta = ti.GetTypeAttr()
        if ta.typekind == pythoncom.TKIND_ENUM:
            enum = ti.GetDocumentation(-1)[0]
            vals = {}
            for j in range(ta.cVars):
                vd = ti.GetVarDesc(j)
                vals[ti.GetNames(vd.memid)[0]] = vd.value
            enums[enum] = vals
    return enums, major


def main(argv):
    if not argv:
        print(__doc__)
        return 1
    enums, major = load_enums()
    if argv[0] == "--json":
        json.dump(enums, sys.stdout, indent=1, sort_keys=True)
        return 0
    print(f"# swconst, SOLIDWORKS type library major version {major}")
    hits = 0
    for pat in argv:
        low = pat.lower()
        for enum, vals in sorted(enums.items()):
            if enum.lower() == low:
                print(f"\n{enum}")
                for k, v in sorted(vals.items(), key=lambda kv: kv[1]):
                    print(f"  {k} = {v}")
                hits += len(vals)
                continue
            for k, v in sorted(vals.items(), key=lambda kv: kv[1]):
                if low in k.lower():
                    print(f"  {k} = {v}    ({enum})")
                    hits += 1
    if not hits:
        print("no match")
    return 0 if hits else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
