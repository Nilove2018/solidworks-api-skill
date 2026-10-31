"""Print the exact argument list of SOLIDWORKS API methods from the installed type library.

Never guess an API call - ask the type library:

    python tools/sw_signature.py IFeatureManager.FeatureExtrusion3
    python tools/sw_signature.py IAssemblyDoc.AddMate5 ISketchManager.CreateArc
    python tools/sw_signature.py IAssemblyDoc            # list every member of an interface
    python tools/sw_signature.py --find Chamfer          # search method names in all interfaces

Reads the registered sldworks type library directly (no SOLIDWORKS session needed).
Set SW_TLB_MAJOR to pick a version (33 = SOLIDWORKS 2025, 32 = 2024, 31 = 2023).
"""
import os
import sys

import pythoncom

SLDWORKS_TLB = '{83A33D31-27C5-11CE-BFD4-00400513BB57}'
VT = {2: "short", 3: "long", 4: "float", 5: "double", 8: "string", 9: "object", 11: "bool",
      12: "variant", 13: "unknown", 16: "byte", 17: "byte", 18: "ushort", 19: "ulong", 20: "int64",
      22: "int", 23: "uint", 24: "void", 25: "hresult", 26: "ptr", 27: "array"}
INVOKE = {1: "", 2: "get ", 4: "put ", 8: "putref "}


def load_typelib():
    forced = os.environ.get("SW_TLB_MAJOR")
    for major in ([int(forced)] if forced else range(45, 19, -1)):
        try:
            return pythoncom.LoadRegTypeLib(SLDWORKS_TLB, major, 0, 0), major
        except pythoncom.com_error:
            continue
    sys.exit("No SOLIDWORKS type library registered - is SOLIDWORKS installed?")


def type_name(td, ti):
    """Readable name of a TYPEDESC: an int VARTYPE, (26, inner) pointer or (29, href) user type."""
    if isinstance(td, int):
        name = VT.get(td & 0x0FFF, f"vt{td & 0x0FFF}")
        return name + "[]" if td & 0x2000 else name
    vt, inner = td
    if vt == 26:                                    # VT_PTR: interface pointer or ByRef
        if isinstance(inner, tuple) and inner[0] == 29:
            return type_name(inner, ti)
        return "ByRef " + type_name(inner, ti)
    if vt == 29:                                    # VT_USERDEFINED: resolve the interface name
        try:
            return ti.GetRefTypeInfo(inner).GetDocumentation(-1)[0]
        except pythoncom.com_error:
            return "object"
    if vt == 27:                                    # VT_SAFEARRAY
        return type_name(inner, ti) + "[]"
    return VT.get(vt, f"vt{vt}")


def interfaces(tlb):
    out = {}
    for i in range(tlb.GetTypeInfoCount()):
        ti = tlb.GetTypeInfo(i)
        ta = ti.GetTypeAttr()
        if ta.typekind in (pythoncom.TKIND_DISPATCH, pythoncom.TKIND_INTERFACE):
            out[ti.GetDocumentation(-1)[0]] = ti
    return out


def members(ti):
    ta = ti.GetTypeAttr()
    for j in range(ta.cFuncs):
        fd = ti.GetFuncDesc(j)
        names = ti.GetNames(fd.memid)
        params = [f"{type_name(p[0], ti)} {n}" for p, n in zip(fd.args, names[1:])]
        yield names[0], fd, params


def show(iface, ti, only=None):
    found = False
    seen = set()
    for name, fd, params in members(ti):
        if only and name != only:
            continue
        key = (name, fd.invkind)
        if key in seen:
            continue
        seen.add(key)
        found = True
        doc = ti.GetDocumentation(fd.memid)[1] or ""
        print(f"{INVOKE.get(fd.invkind, '')}{iface}.{name}({', '.join(params)}) -> {type_name(fd.rettype[0], ti)}"
              f"   [{len(params)} args]")
        if doc and only:
            print(f"    {doc}")
    return found


def main(argv):
    if not argv:
        print(__doc__)
        return 1
    tlb, major = load_typelib()
    ifs = interfaces(tlb)
    print(f"# SOLIDWORKS type library major version {major}\n")
    if argv[0] == "--find":
        pat = argv[1].lower()
        for iface, ti in sorted(ifs.items()):
            for name, fd, params in members(ti):
                if pat in name.lower() and fd.invkind == 1:
                    print(f"{iface}.{name}({', '.join(params)})   [{len(params)} args]")
        return 0
    rc = 0
    for arg in argv:
        iface, _, meth = arg.partition(".")
        ti = ifs.get(iface)
        if ti is None:
            print(f"unknown interface {iface!r} (try IModelDoc2, IFeatureManager, ISketchManager, IAssemblyDoc ...)")
            rc = 1
            continue
        if not show(iface, ti, meth or None):
            print(f"{iface} has no member {meth!r} (try --find {meth})")
            rc = 1
        print()
    return rc


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
