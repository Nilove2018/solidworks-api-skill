"""Check the author credit that swlib.stamp_credit wrote into SOLIDWORKS files, PDFs and STEPs.

    python tools/verify_credit.py FILE [FILE ...] [--credit "Name"] [--key PATH]

Per file it reports each layer and an overall verdict:
  INTACT   a signed record is present, its HMAC signature is valid, and every visible author
           field agrees with it (and with --credit, if given)
  ALTERED  a record or a field was changed (bad signature, or fields that disagree)
  MISSING  no signed record at all (stripped, or never stamped)
The NTFS stream (layer 4) is lost on copies to non-NTFS drives, zips and uploads, and when the
file is re-saved outside swlib, so a missing stream alone is reported but is not an alteration.
SOLIDWORKS files are opened read-only (SOLIDWORKS must be installed); a file you already have
open is read in place and left open. Exit code 0 = all INTACT.

WARNING (2026-10-31): checking SOLIDWORKS files from this separate process, right after another
process had stamped them in the same SOLIDWORKS session, was followed by a SOLIDWORKS crash
(fault module slduiu). Until that is diagnosed, SOLIDWORKS files are only checked with
--sldworks, and only with your work saved. STEP and PDF checks don't touch SOLIDWORKS.
"""
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "examples"))
import swlib as L  # noqa: E402


def verdict(records, fields, key, credit):
    """records: signed records found (None = absent); fields: visible author strings."""
    found = [r for r in records if r]
    if not found:
        return "MISSING", []
    bad = [f"bad signature on '{r.get('credit')}'" for r in found if not L.credit_valid(r, key)]
    names = {r.get("credit") for r in found} | {f for f in fields if f}
    if len({r.get("sig") for r in found}) > 1:
        bad.append("records differ")
    if len(names) > 1:
        bad.append(f"author fields disagree: {sorted(names)}")
    if credit and names != {credit}:
        bad.append(f"expected '{credit}'")
    return ("ALTERED", bad) if bad else ("INTACT", [])


def check_sw(sw, path, key, credit):
    was_open = sw.app.GetOpenDocumentByName(path) is not None
    doc = sw.open(path, readonly=True)
    try:
        summary, prop, rec = L.read_credit(sw.app, doc)
    finally:
        if not was_open:
            sw.close_path(path)
    ads = L.read_file_ads(path)
    layers = {"summary Author": summary or "-", "custom property Author": prop or "-",
              "hidden attribute": _rec_state(rec, key), "NTFS stream": _rec_state(ads, key)}
    return layers, verdict([rec, ads], [summary, prop], key, credit)


def check_pdf(path, key, credit):
    import json
    import pymupdf
    with pymupdf.open(path) as pdf:
        meta = pdf.metadata or {}
    rec = None
    kw = meta.get("keywords") or ""
    if kw.startswith("credit:"):
        try:
            rec = json.loads(kw[len("credit:"):])
        except ValueError:
            rec = {"sig": "unreadable"}
    layers = {"Author": meta.get("author") or "-", "Creator": meta.get("creator") or "-",
              "signed record": _rec_state(rec, key)}
    return layers, verdict([rec], [meta.get("author"), meta.get("creator")], key, credit)


def check_step(path, credit):
    with open(path, encoding="latin-1") as fh:
        head = fh.read(4000)
    m = re.search(r"FILE_NAME\s*\(\s*'(?:[^']|'')*'\s*,\s*'(?:[^']|'')*'\s*,\s*\(\s*'((?:[^']|'')*)'", head)
    author = m.group(1).replace("''", "'") if m else ""
    # a STEP header carries no signature: completeness only
    ok = bool(author) and (not credit or author == credit)
    return {"header author": author or "-"}, (("INTACT" if ok else "MISSING" if not author else "ALTERED"),
                                              [] if ok else [f"author '{author}'"])


def _rec_state(rec, key):
    if not rec:
        return "missing"
    return f"'{rec.get('credit')}' {rec.get('date')}, signature {'valid' if L.credit_valid(rec, key) else 'INVALID'}"


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("files", nargs="+")
    ap.add_argument("--credit", help="the author name every layer must carry")
    ap.add_argument("--key", default=L.DEFAULT_KEY_PATH, help="HMAC key file (never share it)")
    ap.add_argument("--sldworks", action="store_true",
                    help="also check SLDPRT/SLDASM/SLDDRW files (see the crash warning in the docstring)")
    a = ap.parse_args()
    key = L.credit_key(a.key, create=False)
    sw, all_ok = None, True
    for f in a.files:
        path = os.path.abspath(f)
        e = os.path.splitext(path)[1].lower()
        if e in (".sldprt", ".sldasm", ".slddrw"):
            if not a.sldworks:
                print(f"{f}: skipped (SOLIDWORKS file; add --sldworks, with your work saved)")
                continue
            sw = sw or L.SwApp()
            layers, (state, why) = check_sw(sw, path, key, a.credit)
        elif e == ".pdf":
            layers, (state, why) = check_pdf(path, key, a.credit)
        elif e in (".step", ".stp"):
            layers, (state, why) = check_step(path, a.credit)
        else:
            print(f"{f}: skipped (unsupported type)")
            continue
        all_ok &= state == "INTACT"
        print(f"{state:8s} {f}" + (f"  ({'; '.join(why)})" if why else ""))
        for k, v in layers.items():
            print(f"         {k:24s} {v}")
    sys.exit(0 if all_ok else 1)


if __name__ == "__main__":
    main()
