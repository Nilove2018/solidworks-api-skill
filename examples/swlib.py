"""swlib — a small SOLIDWORKS API layer for Python (pywin32), used by the example scripts.

Everything is expressed in *model millimetres* in one product frame: SOLIDWORKS'
Front Plane is the XY plane, Top Plane is XZ, Right Plane is YZ.

Design rules of this layer
* Early-bound COM (makepy wrappers of the installed SOLIDWORKS type library; the newest
  installed version is picked, or set SW_TLB_MAJOR, e.g. 33 = SOLIDWORKS 2025).
* Sketch geometry is converted model -> sketch with the sketch's own transform,
  so the same code works on any plane (no hand-written axis swaps).
* Faces and edges are chosen by geometry (normal, position), never by screen picks.
* Every feature gets a readable name; sketches can be fully defined with Fixed
  relations; failures raise immediately with the feature name.

* Documents the script creates and never saves are closed on exit (also after a crash);
  nothing else is ever closed, and files the user already has open are reused.

Importing this module does not start or touch SOLIDWORKS; SwApp() attaches to it.
Tested on SOLIDWORKS 2025 SP1.2 with Python 3.11 and pywin32.

Version 2.1 adds the helpers from the torsion-spring build (configurations, helix + sweep,
combine, exact extents, set_global, angle-limit mates, mate errors, flexible sub-assemblies,
two-pass interference, document tracking). All of them ran early-bound from this file on
SOLIDWORKS 2025 SP1.2 (SolidWorks Test Runs, 2026-10-30). The startup wait in SwApp._start has not
been run yet.

Version 2.2 adds the helpers from the caliper and vent builds (self-checking mates, save by
path, plane through 3 points, entity names, feature groups, ray casting, backup, log). All of
them ran early-bound on SOLIDWORKS 2025 SP1.2 (SolidWorks Test Run, 2026-10-31).

Skill 2.3 adds the author-credit helpers (stamp_credit, read_credit, stamp_file_ads, stamp_step,
stamp_pdf; set SwApp.credit to stamp every save). A part, an assembly, a drawing, a STEP and a
PDF checked INTACT (2026-10-31). Verify in the process that stamped: a second
process reading the files was followed by a SOLIDWORKS crash (cause not proven).

Skill 2.4 adds drawing checks that need no SOLIDWORKS: check_layout (planned zones inside the
border and off the title block), pdf_words and box_overlaps (text on text in the saved PDF).
Tested on a synthetic overlap and on a real A3 drawing PDF (2026-11-01).
"""
from __future__ import annotations

import atexit
import math
import os
import time

import pythoncom
import win32com.client as w32
from win32com.client import VARIANT, gencache

SLDWORKS_TLB = '{83A33D31-27C5-11CE-BFD4-00400513BB57}'
SWCONST_TLB = '{4687F359-55D0-4CD3-B6CF-2EB42C11F989}'


def _load_typelibs():
    """Early-bound wrappers of the newest installed SOLIDWORKS type library
    (2025 = 33, 2024 = 32, 2023 = 31, ...). SW_TLB_MAJOR forces one version."""
    forced = os.environ.get("SW_TLB_MAJOR")
    majors = [int(forced)] if forced else range(45, 19, -1)
    for major in majors:
        try:
            sw = gencache.EnsureModule(SLDWORKS_TLB, 0, major, 0)
            swc = gencache.EnsureModule(SWCONST_TLB, 0, major, 0)
        except pythoncom.com_error:
            continue
        if sw is not None and swc is not None:
            return sw, swc.constants, major
    raise ImportError("No SOLIDWORKS type library found - is SOLIDWORKS installed and registered?")


SW, SWC, TLB_MAJOR = _load_typelibs()

MM = 0.001
AXES = ("x", "y", "z")


class SwError(RuntimeError):
    pass


def darr(vals):
    """Python numbers -> VARIANT(SAFEARRAY of double) for SolidWorks array arguments."""
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(v) for v in vals])


def cast(obj, iface: str):
    """Return `obj` as the early-bound wrapper of interface `iface` (QueryInterface)."""
    if obj is None:
        return None
    cls = getattr(SW, iface)
    o = obj._oleobj_ if hasattr(obj, "_oleobj_") else obj
    try:
        o = o.QueryInterface(cls.CLSID, pythoncom.IID_IDispatch)
    except pythoncom.com_error:
        pass
    return cls(o)


# "Nothing" for an argument typed `variant` (an argument typed as an interface takes None)
NOTHING = VARIANT(pythoncom.VT_DISPATCH, None)


def dispatch_array(objs):
    """COM objects -> VARIANT(SAFEARRAY of IDispatch) for `variant` arguments (bodies, entities)."""
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH,
                   [o._oleobj_ if hasattr(o, "_oleobj_") else o for o in objs])


def com_put(obj, name, *args):
    """Property put through IDispatch. Works early- and late-bound, for indexed properties
    (IEquationMgr.Equation(i)) and for ones makepy marks read-only (IComponent2.Solving)."""
    o = obj._oleobj_
    o.Invoke(o.GetIDsOfNames(name), 0, pythoncom.DISPATCH_PROPERTYPUT, 0, *args)


def features(doc):
    f = doc.FirstFeature()
    while f is not None:
        f = cast(f, "IFeature")
        yield f
        f = f.GetNextFeature()


def newest(doc, type_name):
    """Last feature of a type in the tree. In an assembly the newest feature is NOT
    FeatureByPositionReverse(0): the MateGroup comes after the reference geometry."""
    out = None
    for f in features(doc):
        if f.GetTypeName2() == type_name:
            out = f
    return out


def delete_feature(doc, feat):
    doc.ClearSelection2(True)
    cast(feat, "IFeature").Select2(False, 0)
    cast(doc.Extension, "IModelDocExtension").DeleteSelection2(0)
    doc.ClearSelection2(True)


def extent(bodies, d):
    """Furthest point (mm) of the bodies along unit direction d. Exact, unlike GetBodyBox,
    which is loose on curved bodies."""
    best = -1e9
    for b in bodies:
        res = cast(b, "IBody2").GetExtremePoint(float(d[0]), float(d[1]), float(d[2]))   # (ok, x, y, z)
        x, y, z = res[1:4]
        best = max(best, (x * d[0] + y * d[1] + z * d[2]) / MM)
    return best


DIRECTIONS = {"+X": (1, 0, 0), "-X": (-1, 0, 0), "+Y": (0, 1, 0), "-Y": (0, -1, 0),
              "+Z": (0, 0, 1), "-Z": (0, 0, -1)}


def ray_hits(doc, bodies, starts, dirs, entry_exit=True):
    """Cast rays at bodies: starts in mm, unit directions. Returns
    [(ray_index, x, y, z, nx, ny, nz, hit_type)] in mm. Batch up to ~4000 rays per call.
    Assert each hit distance against what you expect before using it to cut anything."""
    opts = SWC.swRayPtsOptsNORMALS | (SWC.swRayPtsOptsENTRY_EXIT if entry_exit else 0)
    n = doc.RayIntersections(dispatch_array(bodies), darr([c * MM for p in starts for c in p]),
                             darr([c for d in dirs for c in d]), opts, 0.0, 0.0)
    if not n:
        return []
    a = doc.GetRayIntersectionsPoints() or ()
    out = []
    for i in range(0, len(a) - 8, 9):                  # body, ray, type, x, y, z, nx, ny, nz
        _, r, t, x, y, z, nx, ny, nz = a[i:i + 9]
        out.append((int(r), x / MM, y / MM, z / MM, nx, ny, nz, int(t)))
    return out


def backup(path, reason="backup"):
    """Copy a user's file to <folder>/_Backup/<name>_BACKUP_<date-time>_<reason><ext> and
    verify the copy by MD5 before anything touches the original. Returns (copy, md5)."""
    import hashlib
    import shutil
    folder = os.path.join(os.path.dirname(os.path.abspath(path)), "_Backup")
    os.makedirs(folder, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(path))
    dst = os.path.join(folder, f"{stem}_BACKUP_{time.strftime('%Y%m%d-%H%M%S')}_{reason}{ext}")
    shutil.copy2(path, dst)

    def md5(p):
        with open(p, "rb") as fh:
            return hashlib.md5(fh.read()).hexdigest()
    if md5(path) != md5(dst):
        raise SwError(f"backup of {path} does not match the original")
    return dst, md5(dst)


def log(msg):
    """Timestamped progress line, flushed at once (run scripts with python -u, never | tail)."""
    print(f"[{time.strftime('%H:%M:%S')}] {msg}", flush=True)


# --------------------------------------------------------------------------- drawing checks
def check_layout(items, sheet=(0.42, 0.297), margin=0.02, keepouts=()):
    """Raise SwError unless every planned item lies inside the sheet's inner border and outside
    every keep-out. items: {name: (x0, y0, x1, y1)} in sheet metres (y up, origin bottom-left);
    keepouts: [(name, (x0, y0, x1, y1))], e.g. the title block. Call it with the planned zones
    BEFORE the first view is created, and again with the real text boxes after placing."""
    bad = []
    eps = 1e-9                                      # an item exactly on the inner border is inside
    for name, (x0, y0, x1, y1) in items.items():
        if (x0 < margin - eps or y0 < margin - eps or x1 > sheet[0] - margin + eps
                or y1 > sheet[1] - margin + eps):
            bad.append(f"{name} crosses the border")
        for kname, (kx0, ky0, kx1, ky1) in keepouts:
            if x0 < kx1 and kx0 < x1 and y0 < ky1 and ky0 < y1:
                bad.append(f"{name} overlaps {kname}")
    names = list(items)
    for i, a in enumerate(names):
        for b in names[i + 1:]:
            (ax0, ay0, ax1, ay1), (bx0, by0, bx1, by1) = items[a], items[b]
            if ax0 < bx1 and bx0 < ax1 and ay0 < by1 and by0 < ay1:
                bad.append(f"{a} overlaps {b}")
    if bad:
        raise SwError("layout: " + "; ".join(bad))


def pdf_words(pdf_path, page=0):
    """[(text, (x0, y0, x1, y1))] of every word on a PDF page (points, y down), with pymupdf.
    Check that each expected dimension value is present in full: a screenshot can show a digit
    hidden by a leader line while the PDF text is complete, and the reverse."""
    import pymupdf
    with pymupdf.open(pdf_path) as d:
        return [(w[4], tuple(w[:4])) for w in d[page].get_text("words")]


def box_overlaps(words, shrink=0.2):
    """Pairs of words whose boxes intersect (text on text), from pdf_words(). PDF word boxes
    include the font's line gap, so the lines of one multi-line note overlap by ~30 % of the box
    height; each box is trimmed by `shrink` x its height at top and bottom before comparing."""
    def core(b):
        x0, y0, x1, y1 = b
        h = (y1 - y0) * shrink
        return x0, y0 + h, x1, y1 - h
    out = []
    for i, (ta, a) in enumerate(words):
        ax0, ay0, ax1, ay1 = core(a)
        for tb, b in words[i + 1:]:
            bx0, by0, bx1, by1 = core(b)
            if ax0 < bx1 and bx0 < ax1 and ay0 < by1 and by0 < ay1:
                out.append((ta, tb))
    return out


# --------------------------------------------------------------------------- author credit
# Layers: (1) file summary Author, (2) custom property Author, (3) a hidden API attribute with a
# signed record, (4) the same record in an NTFS alternate data stream of the saved file, (5) an
# HMAC-SHA256 signature with a key that lives only on the owner's machine, (6) PDF / STEP author
# fields. Hidden and tamper-detectable, NOT unremovable: anyone can strip it. Check with
# tools/verify_credit.py. (IModelDoc2.IGet3rdPartyStorage returned None from an external script,
# so a 3rd-party stream needs an add-in's save event and isn't used.)
CREDIT_ATTR_DEF = "AuthorCreditDef"
CREDIT_ATTR = "AuthorCredit"
CREDIT_ADS = "author_credit"
CREDIT_FIELDS = ("credit", "uid", "date", "sig")
DEFAULT_KEY_PATH = os.path.join(os.path.expanduser("~"), ".claude", "credit_key.bin")
_credit_def = None


def credit_key(path=DEFAULT_KEY_PATH, create=True):
    """The HMAC key (32 random bytes), created on first use. Never print, log or commit it."""
    if not os.path.exists(path):
        if not create:
            raise SwError(f"no credit key at {path}")
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "wb") as fh:
            fh.write(os.urandom(32))
    with open(path, "rb") as fh:
        return fh.read()


def credit_sign(key, credit, uid, date):
    import hashlib
    import hmac
    return hmac.new(key, f"{credit}|{uid}|{date}".encode("utf-8"), hashlib.sha256).hexdigest()


def credit_record(credit, key, uid=None, date=None):
    import uuid
    uid = uid or uuid.uuid4().hex
    date = date or time.strftime("%Y-%m-%d")
    return {"credit": credit, "uid": uid, "date": date, "sig": credit_sign(key, credit, uid, date)}


def credit_valid(rec, key):
    import hmac
    return bool(rec) and hmac.compare_digest(rec.get("sig", ""),
                                             credit_sign(key, rec.get("credit", ""), rec.get("uid", ""), rec.get("date", "")))


def _credit_attr_def(app):
    """Attribute definition, defined and registered once per process (needed to write AND read)."""
    global _credit_def
    if _credit_def is None:
        d = cast(app.DefineAttribute(CREDIT_ATTR_DEF), "IAttributeDef")
        for p in CREDIT_FIELDS:
            d.AddParameter(p, SWC.swParamTypeString, 0.0, 0)
        d.Register()               # False if this SOLIDWORKS session registered it already
        _credit_def = d
    return _credit_def


def read_credit(app, doc):
    """(summary author, custom property Author, attribute record or None) of an open document."""
    _credit_attr_def(app)
    doc = cast(doc, "IModelDoc2")
    rec = None
    for f in features(doc):
        if f.GetTypeName2() == "Attribute" and f.Name.startswith(CREDIT_ATTR):
            a = cast(f.GetSpecificFeature2(), "IAttribute")
            rec = {k: cast(a.GetParameter(k), "IParameter").GetStringValue() for k in CREDIT_FIELDS}
    cpm = cast(cast(doc.Extension, "IModelDocExtension").CustomPropertyManager(""), "ICustomPropertyManager")
    res = cpm.Get6("Author", False)
    prop = res[1] if isinstance(res, tuple) else ""
    return doc.SummaryInfo(SWC.swSumInfoAuthor), prop, rec


def stamp_credit(app, doc, credit, key=None):
    """Write layers 1-3 into an open document (call it before saving). Returns the record;
    write layer 4 with stamp_file_ads(path, record) after the save."""
    key = key or credit_key()
    doc = cast(doc, "IModelDoc2")
    com_put(doc, "SummaryInfo", SWC.swSumInfoAuthor, credit)
    cpm = cast(cast(doc.Extension, "IModelDocExtension").CustomPropertyManager(""), "ICustomPropertyManager")
    cpm.Add3("Author", SWC.swCustomInfoText, credit, SWC.swCustomPropertyReplaceValue)
    old = read_credit(app, doc)[2]
    rec = credit_record(credit, key, *((old["uid"], old["date"]) if old and old.get("credit") == credit else ()))
    for f in [f for f in features(doc) if f.GetTypeName2() == "Attribute" and f.Name.startswith(CREDIT_ATTR)]:
        cast(f.GetSpecificFeature2(), "IAttribute").Delete(True)
    att = _credit_attr_def(app).CreateInstance5(doc, None, CREDIT_ATTR, 1, SWC.swAllConfiguration)   # 1 = hidden
    if att is None:
        raise SwError("credit attribute was not created")
    att = cast(att, "IAttribute")
    for k in CREDIT_FIELDS:
        cast(att.GetParameter(k), "IParameter").SetStringValue2(rec[k], SWC.swAllConfiguration, "")
    if read_credit(app, doc)[2] != rec:
        raise SwError("credit attribute did not read back")
    return rec


def stamp_file_ads(path, rec):
    """Layer 4: the signed record in an NTFS alternate data stream of the file. Invisible in
    Explorer; lost when the file is copied to a non-NTFS drive, zipped or uploaded."""
    import json
    with open(f"{path}:{CREDIT_ADS}", "w", encoding="utf-8") as fh:
        json.dump(rec, fh)


def read_file_ads(path):
    import json
    try:
        with open(f"{path}:{CREDIT_ADS}", encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return None


def stamp_step(path, credit):
    """Layer 6: author field of the STEP header FILE_NAME(name, time, (author), (org), ...)."""
    import re
    with open(path, encoding="latin-1") as fh:
        text = fh.read()
    q = credit.replace("'", "''")
    new, n = re.subn(r"(FILE_NAME\s*\(\s*'(?:[^']|'')*'\s*,\s*'(?:[^']|'')*'\s*,\s*)\([^)]*\)",
                     lambda m: f"{m.group(1)}('{q}')", text, count=1)
    if n != 1:
        raise SwError(f"no FILE_NAME header in {path}")
    with open(path, "w", encoding="latin-1", newline="") as fh:
        fh.write(new)


def stamp_pdf(path, credit, rec=None):
    """Layer 6: PDF Author / Creator, with the signed record in Keywords."""
    import json
    import pymupdf
    with pymupdf.open(path) as pdf:
        meta = dict(pdf.metadata or {})
        meta.update(author=credit, creator=credit)
        if rec:
            meta["keywords"] = "credit:" + json.dumps(rec, separators=(",", ":"))
        pdf.set_metadata(meta)
        pdf.saveIncr()


# --------------------------------------------------------------------------- app
def cadmaster_instance():
    """Inside the CAD Master add-in (CADMASTER_SW_PID is set), the SOLIDWORKS that hosts the panel,
    found by its ROT moniker "SolidWorks_PID_<pid>". GetActiveObject could return another running
    (e.g. a hidden -Embedding) SOLIDWORKS. Returns None outside CAD Master or if it isn't found."""
    pid = os.environ.get("CADMASTER_SW_PID", "")
    if not pid or pid == "0":
        return None
    rot = pythoncom.GetRunningObjectTable()
    ctx = pythoncom.CreateBindCtx(0)
    for moniker in rot.EnumRunning():
        try:
            if moniker.GetDisplayName(ctx, None) == "SolidWorks_PID_" + pid:
                return rot.GetObject(moniker)
        except pythoncom.com_error:
            pass
    return None


class SwApp:
    def __init__(self, visible: bool = True, timeout: float = 240.0):
        raw = cadmaster_instance()
        if raw is not None:
            self.app = cast(raw, "ISldWorks")
        elif os.environ.get("CADMASTER_SW_PID", "0") != "0":
            # Inside CAD Master never fall back to GetActiveObject/Dispatch: that attaches to, or
            # starts, a hidden -Embedding SOLIDWORKS, which then swallows every later launch.
            raise SwError("The SOLIDWORKS that hosts CAD Master (PID %s) is not running or not "
                          "answering. Ask the user to restart SOLIDWORKS." % os.environ["CADMASTER_SW_PID"])
        else:
            try:
                self.app = cast(w32.GetActiveObject("SldWorks.Application"), "ISldWorks")
                self.app.GetUserPreferenceStringValue(SWC.swDefaultTemplatePart)
            except pythoncom.com_error:
                self.app = self._start(timeout)
        self.app.Visible = visible
        self.math = cast(self.app.GetMathUtility(), "IMathUtility")
        self.part_template = self.app.GetUserPreferenceStringValue(SWC.swDefaultTemplatePart)
        self.asm_template = self.app.GetUserPreferenceStringValue(SWC.swDefaultTemplateAssembly)
        self.created = []                   # titles of unsaved documents this script made
        self.credit = None                  # set to an author name to stamp every save/export
        atexit.register(self.close_created)

    def save_doc(self, doc, path):
        """SaveAs3 to `path` (closing an older copy by path first), with the author credit
        stamped before the save and the file stream written after it when self.credit is set."""
        doc = cast(doc, "IModelDoc2")
        self.close_other(path, doc)
        rec = stamp_credit(self.app, doc, self.credit) if self.credit else None
        ext = cast(doc.Extension, "IModelDocExtension")
        res = ext.SaveAs3(str(path), 0, SWC.swSaveAsOptions_Silent, None, None, 0, 0)
        ok = res[0] if isinstance(res, tuple) else res
        if not ok:
            raise SwError(f"save {path} failed: {res} (2049 = path too long)")
        if rec:
            stamp_file_ads(str(path), rec)
        return path

    def stamp_export(self, path):
        """Author fields of an exported STEP or PDF (no-op without self.credit)."""
        if not self.credit:
            return
        e = os.path.splitext(str(path))[1].lower()
        if e in (".step", ".stp"):
            stamp_step(str(path), self.credit)
        elif e == ".pdf":
            stamp_pdf(str(path), self.credit, credit_record(self.credit, credit_key()))

    def _start(self, timeout):
        """Start SOLIDWORKS (or wait for one that is still loading). While the splash screen
        shows, calls fail with 'The interface is unknown' (-2147023179): retry until it answers."""
        self._launcher = w32.Dispatch("SldWorks.Application")       # keep a reference while it loads
        end = time.time() + timeout
        while True:
            try:
                app = cast(w32.GetActiveObject("SldWorks.Application"), "ISldWorks")
                app.GetUserPreferenceStringValue(SWC.swDefaultTemplatePart)
                return app
            except pythoncom.com_error:
                if time.time() > end:
                    raise SwError(f"SOLIDWORKS did not answer API calls within {timeout:.0f} s")
                time.sleep(3)

    def close_created(self):
        """Close the documents this script created and never saved (also after a crash).
        Nothing else is ever closed."""
        for t in reversed(self.created):
            try:
                self.app.CloseDoc(t)
            except pythoncom.com_error:
                pass
        self.created = []

    def _untrack(self, title):
        if title in self.created:
            self.created.remove(title)

    def open_docs(self):
        out = []
        docs = self.app.GetDocuments() or ()
        for d in docs:
            d = cast(d, "IModelDoc2")
            out.append((d.GetTitle(), d.GetPathName()))
        return out

    def close(self, doc):
        title = doc.GetTitle()
        self._untrack(title)
        self.app.CloseDoc(title)

    def close_path(self, path: str):
        """Close an open document by its full path (only ours)."""
        for d in self.app.GetDocuments() or ():
            d = cast(d, "IModelDoc2")
            if os.path.normcase(d.GetPathName()) == os.path.normcase(path):
                self.app.CloseDoc(d.GetTitle())

    def close_other(self, path: str, keep):
        """Before saving `keep` to `path`: close any OTHER open document at that path (an earlier
        result). Close by path: after SaveAs3 a document's title changes, so its old title
        closes nothing and the next save fails with error 1."""
        d = self.app.GetOpenDocumentByName(str(path))
        if d is not None and cast(d, "IModelDoc2").GetTitle() != keep.GetTitle():
            self.close_path(path)

    def close_project_docs(self, root: str):
        """Close every open document stored under `root` plus untitled ones (assemblies first)."""
        root = os.path.normcase(os.path.abspath(root))
        docs = []
        for d in self.app.GetDocuments() or ():
            d = cast(d, "IModelDoc2")
            path = d.GetPathName()
            if path == "" or os.path.normcase(os.path.abspath(path)).startswith(root):
                docs.append((0 if d.GetType() == SWC.swDocASSEMBLY else 1, d.GetTitle()))
        for _, t in sorted(docs):
            self.app.CloseDoc(t)
        return len(docs)

    def open(self, path: str, readonly=False):
        """Open a file, or reuse it if it is already open (then it is never closed by us)."""
        doc = self.app.GetOpenDocumentByName(str(path))
        if doc is not None:
            return cast(doc, "IModelDoc2")
        ext = os.path.splitext(path)[1].lower()
        t = {".sldprt": SWC.swDocPART, ".sldasm": SWC.swDocASSEMBLY, ".slddrw": SWC.swDocDRAWING}[ext]
        opt = SWC.swOpenDocOptions_Silent | (SWC.swOpenDocOptions_ReadOnly if readonly else 0)
        res = self.app.OpenDoc6(str(path), t, opt, "", 0, 0)
        doc = res[0] if isinstance(res, tuple) else res
        if doc is None:
            raise SwError(f"cannot open {path}: {res}")
        return cast(doc, "IModelDoc2")


# --------------------------------------------------------------------------- part
class Part:
    """A new part document with model-coordinate modelling helpers."""

    def __init__(self, sw: SwApp, title: str = "", doc=None):
        self.sw = sw
        if doc is None:
            doc = sw.app.NewDocument(sw.part_template, 0, 0, 0)
        if doc is None:
            raise SwError("NewDocument failed")
        self.doc = cast(doc, "IModelDoc2")
        self.part = cast(self.doc, "IPartDoc")
        self.ext = cast(self.doc.Extension, "IModelDocExtension")
        self.sm = cast(self.doc.SketchManager, "ISketchManager")
        self.fm = cast(self.doc.FeatureManager, "IFeatureManager")
        self.selm = cast(self.doc.SelectionManager, "ISelectionMgr")
        self.sm.AddToDB = True              # no inferencing / snapping
        self.sm.DisplayWhenAdded = False
        self.title = title
        self._planes = {}                   # (axis, value) -> plane feature name
        self._names = set()
        self._xf = None                     # active sketch model->sketch transform
        self._axis = None
        self._w = None
        self.base = self._base_planes()     # {'z': 'Front Plane', 'y': 'Top Plane', 'x': 'Right Plane'}
        f = self.doc.FirstFeature()
        while f is not None:
            f = cast(f, "IFeature"); self._names.add(f.Name); f = f.GetNextFeature()
        self.log = []
        self._untitled = self.doc.GetTitle()
        if not self.doc.GetPathName():
            sw.created.append(self._untitled)

    # ---------------------------------------------------------------- utilities
    def _base_planes(self):
        names, f = [], self.doc.FirstFeature()
        while f is not None and len(names) < 3:
            f = cast(f, "IFeature")
            if f.GetTypeName2() == "RefPlane":
                names.append(f.Name)
            f = f.GetNextFeature()
        if len(names) < 3:
            raise SwError("template has no 3 reference planes")
        return {"z": names[0], "y": names[1], "x": names[2]}

    def graphics(self, on: bool):
        try:
            v = cast(self.doc.ActiveView, "IModelView")
            v.EnableGraphicsUpdate = on
        except Exception:
            pass
        try:
            self.fm.EnableFeatureTree = on
        except Exception:
            pass

    def uname(self, name: str) -> str:
        base, i = name, 2
        while name in self._names:
            name = f"{base}_{i}"; i += 1
        self._names.add(name)
        return name

    def last_feature(self):
        f = self.doc.FeatureByPositionReverse(0)
        return cast(f, "IFeature") if f is not None else None

    def rename(self, feat, name: str):
        feat = cast(feat, "IFeature")
        feat.Name = self.uname(name)
        return feat

    def select(self, name: str, typ: str, append=False, mark=0):
        ok = self.ext.SelectByID2(name, typ, 0, 0, 0, append, mark, None, 0)
        if not ok:
            raise SwError(f"cannot select {typ} '{name}'")

    def clear(self):
        self.doc.ClearSelection2(True)

    # ------------------------------------------------------------------ planes
    def plane(self, axis: str, value: float) -> str:
        """Reference plane normal to `axis` at coordinate `value` (mm). Cached."""
        value = round(float(value), 4)
        if abs(value) < 1e-9:
            return self.base[axis]
        key = (axis, value)
        if key in self._planes:
            return self._planes[key]
        flags = SWC.swRefPlaneReferenceConstraint_Distance
        if value < 0:
            flags |= SWC.swRefPlaneReferenceConstraint_OptionFlip
        # the side SOLIDWORKS picks isn't guaranteed: verify with the plane transform,
        # and on the wrong side delete it and retry with OptionFlip toggled
        for _ in range(2):
            self.clear()
            self.select(self.base[axis], "PLANE")
            if self.fm.InsertRefPlane(flags, abs(value) * MM, 0, 0, 0, 0) is None:
                raise SwError(f"plane {axis}={value} failed")
            feat = self.last_feature()
            got = self.plane_origin(feat)[0][AXES.index(axis)]
            if abs(got - value) <= 1e-4:
                break
            delete_feature(self.doc, feat)
            flags ^= SWC.swRefPlaneReferenceConstraint_OptionFlip
        else:
            raise SwError(f"plane {axis}={value} landed at {got}")
        name = self.rename(feat, f"PLN_{axis.upper()}{value:+.2f}").Name
        self._planes[key] = name
        self.clear()
        return name

    @staticmethod
    def plane_origin(feat):
        """Origin (mm) and unit normal of a reference plane feature."""
        rp = cast(cast(feat, "IFeature").GetSpecificFeature2(), "IRefPlane")
        a = cast(rp.Transform, "IMathTransform").ArrayData
        return [v / MM for v in a[9:12]], [a[6], a[7], a[8]]

    def feature(self, name: str):
        """Feature by name (FeatureByName lives on IPartDoc, not IModelDoc2). Look features up
        again after every configuration switch: old handles go stale."""
        f = self.part.FeatureByName(name)
        if f is None:
            raise SwError(f"no feature '{name}'")
        return cast(f, "IFeature")

    def plane_3pts(self, name: str, p1, p2, p3):
        """Reference plane through three model points (mm), via a hidden helper 3D sketch.
        For oblique walls and tilted skins. Checked: all three points lie on the result."""
        self.sm.Insert3DSketch(True)
        pts = [self.sm.CreatePoint(x * MM, y * MM, z * MM) for x, y, z in (p1, p2, p3)]
        self.sm.Insert3DSketch(True)
        sk = self.rename(self.last_feature(), name + "_PTS")
        self.clear()
        for i, pt in enumerate(pts):
            sd = cast(self.selm.CreateSelectData(), "ISelectData")
            sd.Mark = i                                     # marks 0, 1, 2
            cast(pt, "ISketchPoint").Select4(i > 0, sd)
        c = SWC.swRefPlaneReferenceConstraint_Coincident
        if self.fm.InsertRefPlane(c, 0, c, 0, c, 0) is None:
            raise SwError(f"plane {name} through 3 points failed")
        f = self.last_feature()
        org, n = self.plane_origin(f)
        off = max(abs(sum((p[i] - org[i]) * n[i] for i in range(3))) for p in (p1, p2, p3))
        if off > 1e-3:
            raise SwError(f"plane {name}: a point is {off:.4f} mm off the plane")
        self.clear()
        sk.Select2(False, 0)
        self.doc.BlankSketch()
        self.clear()
        return self.rename(f, name)

    def name_entity(self, entity, name: str):
        """Give a face/edge/vertex a name while building, so drawing and assembly scripts find it
        directly with entity() instead of walking thousands of faces."""
        if not self.part.SetEntityName(entity, name):
            raise SwError(f"could not name entity '{name}'")

    def entity(self, name: str, kind: str = "face"):
        t, iface = {"face": (SWC.swSelFACES, "IFace2"), "edge": (SWC.swSelEDGES, "IEdge"),
                    "vertex": (SWC.swSelVERTICES, "IVertex")}[kind]
        e = self.part.GetEntityByName(name, t)
        if e is None:
            raise SwError(f"no {kind} named '{name}'")
        return cast(e, iface)

    def folder(self, folder_name: str, names):
        """Put features in one tree folder (a feature group that can be removed in one go)."""
        self.clear()
        for n in names:
            self.feature(n).Select2(True, 0)
        f = self.fm.InsertFeatureTreeFolder2(SWC.swFeatureTreeFolder_Containing)
        if f is None:
            raise SwError(f"folder '{folder_name}' failed")
        self.clear()
        return self.rename(f, folder_name)

    def delete_group(self, prefix: str):
        """Delete every feature whose name starts with `prefix`, with its absorbed sketches and
        children (a revision = delete the group, rebuild it). Returns the number deleted."""
        names = [f.Name for f in features(self.doc) if f.Name.startswith(prefix)]
        if not names:
            return 0
        self.clear()
        for n in names:
            self.feature(n).Select2(True, 0)
        self.ext.DeleteSelection2(SWC.swDelete_Absorbed | SWC.swDelete_Children)
        self.clear()
        self.doc.ForceRebuild3(False)
        left = [f.Name for f in features(self.doc) if f.Name.startswith(prefix)]
        if left:
            raise SwError(f"delete_group {prefix}: still there {left}")
        return len(names)

    def axis(self, name: str, axis: str, c1: float, c2: float) -> str:
        """Named reference axis parallel to `axis` through (c1, c2) on the other two axes,
        in profile order: z-axis -> (x, y); y-axis -> (x, z); x-axis -> (y, z)."""
        a, b = {"z": ("x", "y"), "y": ("x", "z"), "x": ("y", "z")}[axis]
        pa, pb = self.plane(a, c1), self.plane(b, c2)
        self.clear()
        self.select(pa, "PLANE")
        self.select(pb, "PLANE", append=True)
        if not self.doc.InsertAxis2(True):
            raise SwError(f"axis {name} failed")
        f = self.last_feature()
        if f is None or f.GetTypeName2() != "RefAxis":
            raise SwError(f"axis {name} not created")
        self.clear()
        return self.rename(f, name).Name

    def named_plane(self, name: str, axis: str, value: float) -> str:
        """A reference plane with a readable name (separate from the cached work planes)."""
        self.clear()
        self.select(self.base[axis], "PLANE")
        flags = SWC.swRefPlaneReferenceConstraint_Distance
        if value < 0:
            flags |= SWC.swRefPlaneReferenceConstraint_OptionFlip
        if abs(value) < 1e-9:
            flags = SWC.swRefPlaneReferenceConstraint_Coincident
        if self.fm.InsertRefPlane(flags, abs(value) * MM, 0, 0, 0, 0) is None:
            raise SwError(f"plane {name} failed")
        f = self.last_feature()
        self.clear()
        return self.rename(f, name).Name

    def hide_refs(self):
        try:
            self.doc.SetUserPreferenceToggle(SWC.swViewDisplayHideAllTypes, True)
        except Exception:
            pass

    # ------------------------------------------------------------------ sketch
    def _uvw_to_xyz(self, u, v, w):
        a = self._axis
        if a == "z":
            return u, v, w
        if a == "y":
            return u, w, v
        return w, u, v  # 'x': (u,v) = (y,z)

    def _s(self, u, v):
        """Profile point (u,v) on the active plane -> sketch coordinates (m)."""
        x, y, z = self._uvw_to_xyz(u, v, self._w)
        p = cast(self.sw.math.CreatePoint(darr([x * MM, y * MM, z * MM])), "IMathPoint")
        p = cast(p.MultiplyTransform(self._xf), "IMathPoint")
        a = p.ArrayData
        return a[0], a[1], 0.0

    def sketch(self, axis: str, w: float, name: str = "SK"):
        """Open a sketch on the plane normal to `axis` at `w`. Profile coords:
        z-plane (u,v)=(x,y); y-plane (u,v)=(x,z); x-plane (u,v)=(y,z)."""
        pl = self.plane(axis, w)
        self.clear()
        self.select(pl, "PLANE")
        self.sm.InsertSketch(True)
        sk = cast(self.sm.ActiveSketch, "ISketch")
        if sk is None:
            raise SwError("sketch did not open")
        self._xf = cast(sk.ModelToSketchTransform, "IMathTransform")
        self._axis, self._w = axis, float(w)
        # sign of model +axis in sketch space -> default extrude direction
        e = [0.0, 0.0, 0.0]; e[AXES.index(axis)] = 1.0
        p0 = cast(self.sw.math.CreatePoint(darr([0.0, 0.0, 0.0])), "IMathPoint")
        p1 = cast(self.sw.math.CreatePoint(darr(e)), "IMathPoint")
        z0 = cast(p0.MultiplyTransform(self._xf), "IMathPoint").ArrayData[2]
        z1 = cast(p1.MultiplyTransform(self._xf), "IMathPoint").ArrayData[2]
        self._normal_sign = 1 if (z1 - z0) > 0 else -1
        self._sk_name = name
        return sk

    # profile primitives (all in profile mm) ----------------------------------
    def line(self, u0, v0, u1, v1, construction=False):
        a, b = self._s(u0, v0), self._s(u1, v1)
        seg = self.sm.CreateLine(*a, *b)
        if seg is None:
            raise SwError("line failed")
        if construction:
            cast(seg, "ISketchSegment").ConstructionGeometry = True
        return seg

    def centerline(self, u0, v0, u1, v1):
        a, b = self._s(u0, v0), self._s(u1, v1)
        seg = self.sm.CreateCenterLine(*a, *b)
        if seg is None:
            raise SwError("centerline failed")
        return seg

    def arc(self, uc, vc, u0, v0, u1, v1, ccw_in_profile=True):
        """Arc around (uc,vc) from p0 to p1, CCW in *profile* space when ccw_in_profile."""
        c, a, b = self._s(uc, vc), self._s(u0, v0), self._s(u1, v1)
        # orientation of profile frame in sketch frame
        pu = self._s(uc + 1, vc); pv = self._s(uc, vc + 1)
        det = (pu[0] - c[0]) * (pv[1] - c[1]) - (pu[1] - c[1]) * (pv[0] - c[0])
        ccw = ccw_in_profile if det > 0 else not ccw_in_profile
        seg = self.sm.CreateArc(*c, *a, *b, 1 if ccw else -1)
        if seg is None:
            raise SwError("arc failed")
        return seg

    def circle(self, uc, vc, r):
        c = self._s(uc, vc)
        seg = self.sm.CreateCircleByRadius(*c, r * MM)
        if seg is None:
            raise SwError("circle failed")
        return seg

    def rect(self, u0, v0, u1, v1, r=0.0):
        u0, u1 = sorted((u0, u1)); v0, v1 = sorted((v0, v1))
        r = min(r, (u1 - u0) / 2, (v1 - v0) / 2)
        if r <= 1e-6:
            for a, b in (((u0, v0), (u1, v0)), ((u1, v0), (u1, v1)), ((u1, v1), (u0, v1)), ((u0, v1), (u0, v0))):
                self.line(*a, *b)
            return
        full_u = abs((u1 - u0) / 2 - r) < 1e-6
        full_v = abs((v1 - v0) / 2 - r) < 1e-6
        if not full_u:
            self.line(u0 + r, v0, u1 - r, v0)
            self.line(u1 - r, v1, u0 + r, v1)
        if not full_v:
            self.line(u1, v0 + r, u1, v1 - r)
            self.line(u0, v1 - r, u0, v0 + r)
        self.arc(u1 - r, v0 + r, u1 - r, v0, u1, v0 + r)
        self.arc(u1 - r, v1 - r, u1, v1 - r, u1 - r, v1)
        self.arc(u0 + r, v1 - r, u0 + r, v1, u0, v1 - r)
        self.arc(u0 + r, v0 + r, u0, v0 + r, u0 + r, v0)

    def rect_c(self, uc, vc, du, dv, r=0.0):
        self.rect(uc - du / 2, vc - dv / 2, uc + du / 2, vc + dv / 2, r)

    def slot(self, uc, vc, du, dv):
        """Obround (stadium) centred at (uc,vc), overall du x dv."""
        self.rect_c(uc, vc, du, dv, min(du, dv) / 2)

    def polygon(self, pts, close=True):
        n = len(pts)
        for i in range(n if close else n - 1):
            a, b = pts[i], pts[(i + 1) % n]
            self.line(a[0], a[1], b[0], b[1])

    def text(self, u, v, txt, height=2.5, width_factor=1.0, spacing=1.0, font="Arial", bold=True):
        """Sketch text whose baseline-left starts at (u,v)."""
        p = self._s(u, v)
        st = self.doc.InsertSketchText(p[0], p[1], 0.0, txt, 0, 0, 0, int(width_factor * 100), int(spacing * 100))
        if st is None:
            raise SwError(f"text '{txt}' failed")
        st = cast(st, "ISketchText")
        tf = cast(st.GetTextFormat(), "ITextFormat")
        tf.CharHeight = height * MM
        tf.TypeFaceName = font
        tf.Bold = bold
        st.SetTextFormat(False, tf)
        return st

    def _dir_s(self, d):
        """Model direction vector -> 2D sketch direction (unit)."""
        p0 = cast(self.sw.math.CreatePoint(darr([0.0, 0.0, 0.0])), "IMathPoint")
        p1 = cast(self.sw.math.CreatePoint(darr([d[0] * MM, d[1] * MM, d[2] * MM])), "IMathPoint")
        a = cast(p0.MultiplyTransform(self._xf), "IMathPoint").ArrayData
        b = cast(p1.MultiplyTransform(self._xf), "IMathPoint").ArrayData
        v = (b[0] - a[0], b[1] - a[1]); n = math.hypot(*v)
        return v[0] / n, v[1] / n

    def label(self, uc, vc, txt, height, right, up, font="Arial", bold=True, width_factor=1.0):
        """Text centred at profile (uc, vc) that reads correctly for a viewer outside the face:
        `right` = reading direction and `up` = glyph up, both as model unit vectors."""
        rs, us = self._dir_s(right), self._dir_s(up)
        mirror = (rs[0] * us[1] - rs[1] * us[0]) < 0
        base = (-rs[0], -rs[1]) if mirror else rs
        esc = math.atan2(base[1], base[0])
        w = 0.66 * height * width_factor * len(txt)
        cx, cy, cz = self._uvw_to_xyz(uc, vc, self._w)
        sx = cx - right[0] * w / 2 - up[0] * height / 2
        sy = cy - right[1] * w / 2 - up[1] * height / 2
        sz = cz - right[2] * w / 2 - up[2] * height / 2
        p = cast(self.sw.math.CreatePoint(darr([sx * MM, sy * MM, sz * MM])), "IMathPoint")
        a = cast(p.MultiplyTransform(self._xf), "IMathPoint").ArrayData
        st = self.doc.InsertSketchText(a[0], a[1], 0.0, txt, 0, 0, 1 if mirror else 0,
                                       int(width_factor * 100), 100)
        if st is None:
            raise SwError(f"label '{txt}' failed")
        st = cast(st, "ISketchText")
        tf = cast(st.GetTextFormat(), "ITextFormat")
        tf.CharHeight = height * MM
        tf.TypeFaceName = font
        tf.Bold = bold
        tf.Escapement = esc
        st.SetTextFormat(False, tf)
        return st

    def _xyz_to_uv(self, x, y, z):
        a = self._axis
        if a == "z":
            return x, y
        if a == "y":
            return x, z
        return y, z

    def text_outline(self, center, txt, height, right, up, bold=True, family="DejaVu Sans", tol=0.06):
        """Glyph outlines (matplotlib font engine) drawn as sketch lines on the active plane.
        center = model point (x,y,z) of the text centre; right/up = model unit vectors seen by
        the reader. Deterministic on any plane (no SOLIDWORKS text flags)."""
        from matplotlib.font_manager import FontProperties
        from matplotlib.textpath import TextPath
        tp = TextPath((0, 0), txt, size=1.0, prop=FontProperties(family=family, weight="bold" if bold else "normal"))
        polys = [pp for pp in tp.to_polygons() if len(pp) >= 3]
        xs = [p[0] for pp in polys for p in pp]; ys = [p[1] for pp in polys for p in pp]
        x0, x1, y0, y1 = min(xs), max(xs), min(ys), max(ys)
        s = height / (y1 - y0) if txt.strip() else 1.0
        w = (x1 - x0) * s
        n = 0
        for pp in polys:
            pts = []
            for gx, gy in pp:
                lx, ly = (gx - x0) * s - w / 2, (gy - y0) * s - height / 2
                m = [center[i] + right[i] * lx + up[i] * ly for i in range(3)]
                u, v = self._xyz_to_uv(*m)
                if not pts or math.hypot(u - pts[-1][0], v - pts[-1][1]) > tol:
                    pts.append((u, v))
            if len(pts) > 2 and math.hypot(pts[0][0] - pts[-1][0], pts[0][1] - pts[-1][1]) <= tol:
                pts.pop()
            if len(pts) >= 3:
                self.polygon(pts); n += len(pts)
        return w, n

    def fix_all(self):
        """Fully define the active sketch by fixing every segment."""
        sk = cast(self.sm.ActiveSketch, "ISketch")
        segs = sk.GetSketchSegments() or ()
        self.clear()
        n = 0
        for s in segs:
            s = cast(s, "ISketchSegment")
            if s.Select4(True, None):
                n += 1
        if n:
            self.doc.SketchAddConstraints("sgFIXED")
        self.clear()

    def end_sketch(self, fix=True):
        if fix:
            self.fix_all()
        self.sm.InsertSketch(True)
        feat = self.last_feature()
        if feat is None or feat.GetTypeName2() not in ("ProfileFeature", "3DProfileFeature"):
            raise SwError(f"sketch {self._sk_name} did not close ({feat.GetTypeName2() if feat else None})")
        feat = self.rename(feat, self._sk_name)
        return feat.Name

    # ---------------------------------------------------------------- features
    def _select_sketch(self, skname):
        self.clear()
        self.select(skname, "SKETCH")

    def extrude(self, name, w1, cut=False, draft_deg=0.0, draft_out=False, both=False,
                through_all=False, fix=True, merge=True):
        """Close the active sketch and extrude it from its plane (w0) to w1.
        `both`: symmetric through-all (cut) — ignores w1."""
        w0, sign = self._w, self._normal_sign
        skname = self.end_sketch(fix)
        self._select_sketch(skname)
        depth = abs(w1 - w0) * MM
        want = 1 if w1 > w0 else -1
        # Measured on SW2025: boss default = +sketch normal, cut default = -sketch normal.
        # 'Dir' reverses the direction; 'Flip' flips the side to cut (must stay False).
        default = sign if not cut else -sign
        rev = (want != default)
        dchk = draft_deg > 0
        dang = math.radians(draft_deg)
        if cut:
            if both:
                f = self.fm.FeatureCut4(False, False, False, SWC.swEndCondThroughAll, SWC.swEndCondThroughAll,
                                        0.01, 0.01, False, False, False, False, 0, 0, False, False, False, False,
                                        False, True, True, False, False, False, 0, 0, False, False)
            else:
                t1 = SWC.swEndCondThroughAll if through_all else SWC.swEndCondBlind
                f = self.fm.FeatureCut4(True, False, rev, t1, 0, depth, 0, dchk, False, draft_out, False,
                                        dang, 0, False, False, False, False, False, True, True, False, False,
                                        False, 0, 0, False, False)
        else:
            f = self.fm.FeatureExtrusion3(True, False, rev, SWC.swEndCondBlind, 0, depth, 0, dchk, False,
                                          draft_out, False, dang, 0, False, False, False, False, merge, True,
                                          True, 0, 0, False)
        if f is None:
            raise SwError(f"{'cut' if cut else 'extrude'} '{name}' failed (sketch {skname})")
        f = self.rename(f, name)
        self.log.append(f.Name)
        self.clear()
        return f

    def revolve(self, name, cut=False, angle=360.0, fix=True, merge=True):
        """Revolve the active sketch about its (single) centerline."""
        skname = self.end_sketch(fix)
        self._select_sketch(skname)
        f = self.fm.FeatureRevolve2(True, True, False, cut, False, False, SWC.swEndCondBlind, 0,
                                    math.radians(angle), 0, False, False, 0, 0, 0, 0, 0, merge, True, True)
        if f is None:
            raise SwError(f"revolve '{name}' failed")
        f = self.rename(f, name)
        self.log.append(f.Name)
        self.clear()
        return f

    def helix(self, sketch_name, name, height, pitch, revolutions, left_hand=False):
        """Helix on a closed circle sketch (mm). Measured on SW 2025 with the circle on a plane
        parallel to Top: Clockwised=True -> left-hand, advancing +Y, starting on the circle at
        +Z. Feature dimensions: D3 = height, D5 = revolutions. Check with helix_hand()."""
        self._select_sketch(sketch_name)
        self.doc.InsertHelix(False, left_hand, False, False, SWC.swHelixDefinedByHeightAndRevolution,
                             height * MM, pitch * MM, revolutions, 0, 0)
        f = self.last_feature()
        if f is None or f.GetTypeName2() != "Helix":
            raise SwError(f"helix '{name}' failed")
        self.clear()
        return self.rename(f, name)

    @staticmethod
    def helix_hand(helix_feat, axis=(0.0, 1.0, 0.0)):
        """('RH' | 'LH', start mm, end mm) from the helix curve itself (axis through the origin)."""
        rc = cast(cast(helix_feat, "IFeature").GetSpecificFeature2(), "IReferenceCurve")
        edge = cast(rc.GetFirstSegment(), "IEdge")
        cp = cast(edge.GetCurveParams3(), "ICurveParamData")
        s, e = [v / MM for v in cp.StartPoint], [v / MM for v in cp.EndPoint]
        t = cast(edge.GetCurve(), "ICurve").Evaluate2(cp.UMinValue, 1)[3:6]
        # (start x tangent) . axis has the sign of the turning sense; times the travel along the axis
        turn = sum(c * a for c, a in zip((s[1] * t[2] - s[2] * t[1], s[2] * t[0] - s[0] * t[2],
                                          s[0] * t[1] - s[1] * t[0]), axis))
        travel = sum((ei - si) * a for si, ei, a in zip(s, e, axis))
        return ("RH" if turn * travel > 0 else "LH"), s, e

    def sweep_circular(self, path_name, name, diameter, merge=True):
        """Circular-profile sweep along a path feature (helix or sketch). merge=False makes a
        separate body (join it later with combine_add). The sweep must OVERLAP the body
        (0.2-0.5 mm) or keep a clear gap: a sweep that only touches it fails or makes 2 bodies."""
        self.clear()
        self.feature(path_name).Select2(False, 4)            # sweep path = mark 4
        sd = cast(self.fm.CreateDefinition(SWC.swFmSweep), "ISweepFeatureData")
        sd.CircularProfile = True
        sd.CircularProfileDiameter = diameter * MM
        sd.Merge = merge
        f = self.fm.CreateFeature(sd)
        if f is None:
            raise SwError(f"sweep '{name}' failed (touching instead of overlapping?)")
        self.clear()
        f = self.rename(f, name)
        self.log.append(f.Name)
        return f

    def combine_add(self, name):
        """Join all solid bodies with a Combine (Add) feature, e.g. at the end of the tree so
        configurations can suppress the join instead of the joined feature."""
        bodies = self.bodies()
        self.clear()
        f = self.fm.InsertCombineFeature(SWC.SWBODYADD, None, dispatch_array(bodies))
        if f is None or len(self.bodies()) != 1:
            raise SwError(f"combine '{name}' failed ({len(self.bodies())} bodies)")
        return self.rename(f, name)

    # ------------------------------------------------------------ topology
    def bodies(self):
        bs = self.part.GetBodies2(SWC.swSolidBody, True) or ()
        return [cast(b, "IBody2") for b in bs]

    def faces(self):
        for b in self.bodies():
            for f in b.GetFaces() or ():
                yield cast(f, "IFace2")

    @staticmethod
    def face_info(f):
        box = [v / MM for v in f.GetBox()]
        surf = cast(f.GetSurface(), "ISurface")
        planar = bool(surf.IsPlane())
        n = tuple(f.Normal) if planar else (0.0, 0.0, 0.0)
        return planar, n, box

    def find_faces(self, normal=None, at=None, axis=None, within=None, planar=True, tol=1e-3):
        """Faces with unit `normal` (tuple) whose plane coordinate on `axis` equals `at`;
        optional `within` = (u0,v0,u1,v1) box the face bbox must lie inside (profile coords)."""
        out = []
        for f in self.faces():
            pl, n, box = self.face_info(f)
            if planar and not pl:
                continue
            if normal is not None and sum((a - b) ** 2 for a, b in zip(n, normal)) > 1e-6:
                continue
            if axis is not None and at is not None:
                i = AXES.index(axis)
                if abs(box[i] - at) > tol or abs(box[i + 3] - at) > tol:
                    continue
            if within is not None:
                lo, hi = box[:3], box[3:]
                ok = True
                for i, (a, b) in enumerate(zip(lo, hi)):
                    if within[i] is not None and not (within[i][0] - tol <= a and b <= within[i][1] + tol):
                        ok = False
                if not ok:
                    continue
            out.append(f)
        return out

    def select_entities(self, ents, mark=0, append=False):
        sd = cast(self.selm.CreateSelectData(), "ISelectData")
        sd.Mark = mark
        n = 0
        for i, e in enumerate(ents):
            e = cast(e, "IEntity")
            if e.Select4(append or i > 0, sd):
                n += 1
        return n

    def edges_of(self, faces):
        seen, out = set(), []
        for f in faces:
            for e in cast(f, "IFace2").GetEdges() or ():
                e = cast(e, "IEdge")
                key = id(e)
                out.append(e)
        return out

    @staticmethod
    def edge_box(e):
        cp = cast(e.GetCurveParams3(), "ICurveParamData")
        s, t = [v / MM for v in cp.StartPoint], [v / MM for v in cp.EndPoint]
        return s, t

    # ------------------------------------------------------------ operations
    def shell(self, name, faces, t):
        self.clear()
        if self.select_entities(faces, mark=1) != len(faces):
            raise SwError(f"shell '{name}': face selection failed")
        self.doc.InsertFeatureShell(t * MM, False)
        f = self.last_feature()
        if f is None or f.GetTypeName2() != "Shell":
            raise SwError(f"shell '{name}' failed")
        f = self.rename(f, name)
        self.clear()
        return f

    def fillet(self, name, ents, r):
        self.clear()
        if self.select_entities(ents, mark=1) == 0:
            raise SwError(f"fillet '{name}': nothing selected")
        f = self.fm.FeatureFillet3(195, r * MM, 0, 0, 0, 0, 0, None, None, None, None, None, None, None)
        if f is None:
            raise SwError(f"fillet '{name}' failed")
        f = self.rename(f, name)
        self.clear()
        return f

    def chamfer(self, name, ents, d, angle=45.0):
        self.clear()
        if self.select_entities(ents, mark=1) == 0:
            raise SwError(f"chamfer '{name}': nothing selected")
        f = self.fm.InsertFeatureChamfer(4, SWC.swChamferAngleDistance, d * MM, math.radians(angle), 0, 0, 0, 0)
        if f is None:
            raise SwError(f"chamfer '{name}' failed")
        f = self.rename(f, name)
        self.clear()
        return f

    def outer_loop_edges(self, normal, at, axis, rect, r, tol=0.01):
        """Edges of the planar face(s) at `at` whose end points lie on the rounded-rectangle
        contour rect=(x0,y0,x1,y1) with corner radius r (projected on the face plane)."""
        x0, y0, x1, y1 = rect

        def on_contour(p):
            u, v = p[0], p[1]
            cu = min(max(u, x0 + r), x1 - r); cv = min(max(v, y0 + r), y1 - r)
            if (x0 + r - tol <= u <= x1 - r + tol) or (y0 + r - tol <= v <= y1 - r + tol):
                d = min(abs(u - x0), abs(u - x1), abs(v - y0), abs(v - y1))
            else:
                d = abs(math.hypot(u - cu, v - cv) - r)
            return d < tol

        out = []
        for f in self.find_faces(normal=normal, at=at, axis=axis):
            for e in f.GetEdges() or ():
                e = cast(e, "IEdge")
                s, t = self.edge_box(e)
                if on_contour(s) and on_contour(t):
                    out.append(e)
        return out

    # ------------------------------------------------------------ equations
    def globals(self, values: dict):
        eq = cast(self.doc.GetEquationMgr(), "IEquationMgr")
        for k, v in values.items():
            eq.Add2(-1, f'"{k}" = {v}', True)
        return eq

    def equation(self, text: str):
        eq = cast(self.doc.GetEquationMgr(), "IEquationMgr")
        return eq.Add2(-1, text, True)

    def set_global(self, name: str, value: float):
        """Change an existing global variable in place (delete + Add2 fails), then rebuild.
        Equation(i) is an indexed property, so it is set through com_put."""
        eq = cast(self.doc.GetEquationMgr(), "IEquationMgr")
        want = f'"{name}"={value:g}'
        for i in range(eq.GetCount()):
            if eq.Equation(i).replace(" ", "").startswith(f'"{name}"='):
                com_put(eq, "Equation", i, f'"{name}" = {value:g}')
                if eq.Equation(i).replace(" ", "") != want:
                    raise SwError(f"could not set {name}: {eq.Equation(i)}")
                break
        else:
            raise SwError(f"no global variable '{name}'")
        eq.EvaluateAll()
        self.doc.ForceRebuild3(False)

    # ------------------------------------------------------------ configurations
    def configurations(self):
        return list(self.doc.GetConfigurationNames() or ())

    def active_configuration(self):
        return cast(self.doc.GetActiveConfiguration(), "IConfiguration").Name

    def add_configuration(self, name: str, comment: str = ""):
        if self.doc.AddConfiguration3(name, comment, "", 0) is None:
            raise SwError(f"configuration '{name}' failed")

    def suppress_in(self, cfg: str, names, suppress=True):
        """Suppress (or unsuppress) features in ONE configuration: activate it, look the
        features up again by name, EditSuppress2. (SetSuppression2 with swSpecifyConfiguration
        returned True and did nothing.) Dependents follow their parents: when unsuppressing
        after a temporary suppress, list the dependents too."""
        self.doc.ShowConfiguration2(cfg)
        self.clear()
        for i, n in enumerate(names):
            self.feature(n).Select2(i > 0, 0)
        ok = self.doc.EditSuppress2() if suppress else self.doc.EditUnsuppress2()
        self.clear()
        self.doc.ForceRebuild3(False)
        if not ok:
            raise SwError(f"{'suppress' if suppress else 'unsuppress'} {list(names)} in {cfg} failed")

    def suppression(self, cfg: str, names):
        """Read back: tuple of bools (True = suppressed) in configuration `cfg`."""
        self.doc.ShowConfiguration2(cfg)
        return tuple(bool(self.feature(n).IsSuppressed2(SWC.swThisConfiguration, None)[0]) for n in names)

    def check_configurations(self):
        """{config: (bodies, volume mm3, What's Wrong)} for every configuration. Run it just
        before saving; the active configuration is restored."""
        active, out = self.active_configuration(), {}
        for cfg in self.configurations():
            self.doc.ShowConfiguration2(cfg)
            self.doc.ForceRebuild3(False)
            out[cfg] = (len(self.bodies()), self.volume(), self.ext.GetWhatsWrongCount())
        self.doc.ShowConfiguration2(active)
        return out

    # ------------------------------------------------------------ appearance
    def color(self, rgb, shininess=0.3, transparency=0.0, specular=0.3, configs=None):
        """Part colour. Appearance is per configuration: configs=None sets the active one,
        configs="all" sets every configuration."""
        r, g, b = rgb
        vals = darr([r, g, b, 1.0, 0.8, specular, shininess, transparency, 0.0])
        if configs is None:
            self.doc.MaterialPropertyValues = vals
            return
        active = self.active_configuration()
        for cfg in (self.configurations() if configs == "all" else configs):
            self.doc.ShowConfiguration2(cfg)
            self.doc.MaterialPropertyValues = vals
        self.doc.ShowConfiguration2(active)

    def material(self, name="PET", db="SOLIDWORKS Materials", fallbacks=()):
        """Set the material and read it back (SetMaterialPropertyName2 fails silently when the
        name isn't in the database). Returns the name actually applied, or '' if none."""
        for n in (name, *fallbacks):
            try:
                self.part.SetMaterialPropertyName2("", db, n)
            except pythoncom.com_error:
                continue
            res = self.part.GetMaterialPropertyName2("")       # early-bound: (name, database)
            got = res[0] if isinstance(res, tuple) else res
            if got:
                return got
        return ""

    def properties(self, props: dict):
        cpm = cast(self.ext.CustomPropertyManager(""), "ICustomPropertyManager")
        for k, v in props.items():
            cpm.Add3(k, SWC.swCustomInfoText, str(v), SWC.swCustomPropertyReplaceValue)

    # ------------------------------------------------------------ results
    def volume(self):
        mp = cast(self.ext.CreateMassProperty(), "IMassProperty")
        return mp.Volume / (MM ** 3)

    def extent(self, d):
        """Exact furthest point (mm) along unit direction d."""
        return extent(self.bodies(), d)

    def extents(self):
        """Exact extents in the 6 axis directions, e.g. {'+X': 3.0, '-X': 3.0, ...} (mm)."""
        return {k: extent(self.bodies(), d) for k, d in DIRECTIONS.items()}

    def bbox(self):
        """Fast but LOOSE on curved bodies (helices, fillets): use extents() for size checks."""
        lo, hi = [1e9] * 3, [-1e9] * 3
        for b in self.bodies():
            bx = [v / MM for v in b.GetBodyBox()]
            for i in range(3):
                lo[i] = min(lo[i], bx[i]); hi[i] = max(hi[i], bx[i + 3])
        return lo, hi

    def rebuild_check(self):
        self.doc.ForceRebuild3(False)
        n = self.ext.GetWhatsWrongCount()
        return n

    def save(self, path, fit=True):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        self.graphics(True)
        self.hide_refs()
        if fit:
            try:
                self.doc.ShowNamedView2("*Isometric", 7)
                self.doc.ViewZoomtofit2()
            except Exception:
                pass
        self.sw.save_doc(self.doc, path)
        self.sw._untrack(self._untitled)    # a saved file is the user's: never closed on exit
        return path                         # (the document's title is now the file name)

    def export(self, path):
        """Save a copy as STEP/STL/PNG etc. (extension decides)."""
        os.makedirs(os.path.dirname(path), exist_ok=True)
        res = self.ext.SaveAs3(path, 0, SWC.swSaveAsOptions_Silent | SWC.swSaveAsOptions_Copy, None, None, 0, 0)
        ok = res[0] if isinstance(res, tuple) else res
        if not ok:
            raise SwError(f"export {path} failed: {res}")
        self.sw.stamp_export(path)
        return path


def set_export_prefs(sw: SwApp):
    a = sw.app
    a.SetUserPreferenceIntegerValue(SWC.swStepAP, 214)
    a.SetUserPreferenceToggle(SWC.swSTLBinaryFormat, True)
    a.SetUserPreferenceIntegerValue(SWC.swExportStlUnits, SWC.swMM)
    a.SetUserPreferenceIntegerValue(SWC.swSTLQuality, SWC.swSTLQuality_Custom)
    a.SetUserPreferenceDoubleValue(SWC.swSTLDeviation, 0.01 * MM)
    a.SetUserPreferenceDoubleValue(SWC.swSTLAngleTolerance, math.radians(5))
    a.SetUserPreferenceToggle(SWC.swSTLDontTranslateToPositive, True)


# --------------------------------------------------------------------------- assembly
class Asm:
    """New assembly document with placement-then-mate helpers (all positions in mm)."""

    def __init__(self, sw: SwApp):
        self.sw = sw
        doc = sw.app.NewDocument(sw.asm_template, 0, 0, 0)
        if doc is None:
            raise SwError("NewDocument (assembly) failed")
        self.doc = cast(doc, "IModelDoc2")
        self.asm = cast(self.doc, "IAssemblyDoc")
        self.ext = cast(self.doc.Extension, "IModelDocExtension")
        self.fm = cast(self.doc.FeatureManager, "IFeatureManager")
        self.title = self.doc.GetTitle()
        sw.created.append(self.title)
        self.nmates = 0
        self.log = []

    def add(self, path: str, rows=((1, 0, 0), (0, 1, 0), (0, 0, 1)), t=(0.0, 0.0, 0.0), config: str = ""):
        """Insert a part/assembly and place it: rows = images of the component's x, y, z axes,
        t = position of its origin (mm), config = configuration to use ("" = the file's active
        one). Returns IComponent2. The placement is kept only until mates solve it elsewhere."""
        path = str(path)
        self.sw.open(path)                  # must be loaded; reuses a document the user has open
        self.sw.app.ActivateDoc3(self.title, False, 0, 0)
        comp = self.asm.AddComponent5(path, SWC.swAddComponentConfigOptions_CurrentSelectedConfig, "", False,
                                      config, t[0] * MM, t[1] * MM, t[2] * MM)
        if comp is None:
            raise SwError(f"AddComponent5 failed for {path}")
        comp = cast(comp, "IComponent2")
        if config and comp.ReferencedConfiguration != config:
            comp.ReferencedConfiguration = config
            if comp.ReferencedConfiguration != config:
                raise SwError(f"{comp.Name2}: configuration {comp.ReferencedConfiguration}, wanted {config}")
        arr = [*rows[0], *rows[1], *rows[2], t[0] * MM, t[1] * MM, t[2] * MM, 1.0, 0.0, 0.0, 0.0]
        comp.Transform2 = cast(self.sw.math.CreateTransform(darr(arr)), "IMathTransform")
        return comp

    def feat(self, comp, name):
        f = comp.FeatureByName(name)
        if f is None:
            raise SwError(f"{comp.Name2} has no feature '{name}'")
        return cast(f, "IFeature")

    def mate(self, e1, e2, mtype=None, align=None, dist=0.0, label="", flip=False,
             lo=None, hi=None, angle=0.0, ang_lo=None, ang_hi=None, name=None, verify=None):
        """AddMate5. Lengths mm, angles degrees. Limit mates: lo/hi (distance) or ang_lo/ang_hi
        (angle; arguments 9-11 of AddMate5, radians). An angle range must not include 0 or 180
        degrees: measure against a datum at 90 degrees instead. `name` renames the new mate.
        verify: a function returning True when the result is right (e.g. the component's
        position). The mate is then rebuilt and checked, and on a wrong result deleted and
        added again with flip toggled, then with the other alignment, in the same run."""
        mtype = SWC.swMateCOINCIDENT if mtype is None else mtype
        align = SWC.swMateAlignCLOSEST if align is None else align
        kw = dict(dist=dist, label=label, lo=lo, hi=hi, angle=angle, ang_lo=ang_lo, ang_hi=ang_hi, name=name)
        if verify is None:
            return self._mate(e1, e2, mtype, align, flip, **kw)
        others = [a for a in (SWC.swMateAlignALIGNED, SWC.swMateAlignANTI_ALIGNED) if a != align]
        tried = []
        for fl, al in [(flip, align), (not flip, align)] + [(f, a) for a in others for f in (flip, not flip)]:
            try:
                self._mate(e1, e2, mtype, al, fl, **kw)
            except SwError:
                tried.append((fl, al, "rejected"))
                continue
            self.doc.ForceRebuild3(False)
            if verify():
                return self.mates()[-1]
            tried.append((fl, al, "wrong result"))
            delete_feature(self.doc, self.mates()[-1])
            self.doc.ForceRebuild3(False)
        raise SwError(f"mate {label or name}: no flip/alignment passed verify ({tried})")

    def _mate(self, e1, e2, mtype, align, flip, dist=0.0, label="", lo=None, hi=None,
              angle=0.0, ang_lo=None, ang_hi=None, name=None):
        self.doc.ClearSelection2(True)
        if not e1.Select2(False, 1) or not e2.Select2(True, 1):
            raise SwError(f"mate {label or name}: selection failed")
        r = math.radians
        res = self.asm.AddMate5(mtype, align, flip, dist * MM,
                                0 if hi is None else hi * MM, 0 if lo is None else lo * MM, 0, 0,
                                r(angle), r(angle if ang_hi is None else ang_hi),
                                r(angle if ang_lo is None else ang_lo), False, False, 0, 0)
        mate = res[0] if isinstance(res, tuple) else res
        err = res[1] if isinstance(res, tuple) else None
        self.doc.ClearSelection2(True)
        if mate is None:
            raise SwError(f"mate {label or name} failed (swAddMateError_e {err})")
        self.nmates += 1
        if name:
            mate = self.mates()[-1]         # the new mate = last sub-feature of the MateGroup
            mate.Name = name
        return mate

    def newest(self, type_name):
        """Newest feature of a type (FeatureByPositionReverse(0) is the MateGroup here)."""
        return newest(self.doc, type_name)

    def mates(self):
        group = self.newest("MateGroup")
        out, sub = [], group.GetFirstSubFeature() if group is not None else None
        while sub is not None:
            sub = cast(sub, "IFeature")
            out.append(sub)
            sub = sub.GetNextSubFeature()
        return out

    def mate_errors(self):
        """Names of mates with an error code after the last rebuild."""
        bad = []
        for m in self.mates():
            res = m.GetErrorCode2()             # early-bound: (code, is_warning)
            if (res[0] if isinstance(res, tuple) else res):
                bad.append(m.Name)
        return bad

    def make_flexible(self, comp):
        """Make Subassembly Flexible. makepy has no setter for Solving, so put it through IDispatch."""
        com_put(comp, "Solving", SWC.swComponentFlexibleSolving)
        self.doc.ForceRebuild3(False)
        if comp.Solving != SWC.swComponentFlexibleSolving:
            raise SwError(f"{comp.Name2} did not become flexible")

    def interference(self, coincidence=False):
        """[(component names, volume mm3)]. Run with coincidence=False (must be empty), then with
        True to list touching contacts, and check each against your intended list."""
        idm = cast(self.asm.InterferenceDetectionManager, "IInterferenceDetectionMgr")
        idm.TreatCoincidenceAsInterference = coincidence
        idm.TreatSubAssembliesAsComponents = True
        idm.IncludeMultibodyPartInterferences = False
        idm.IgnoreHiddenBodies = True
        idm.UseTransform = True
        out = []
        for it in idm.GetInterferences() or ():
            it = cast(it, "IInterference")
            names = sorted(cast(c, "IComponent2").Name2 for c in it.Components or ())
            out.append((names, it.Volume / MM ** 3))
        idm.Done()
        return out

    def comp_box(self, comp):
        b = comp.GetBox(False, False)
        return [v / MM for v in b] if b else None

    def save(self, path):
        self.sw.save_doc(self.doc, path)
        self.sw._untrack(self.title)
        self.title = self.doc.GetTitle()    # the title changes on SaveAs3
