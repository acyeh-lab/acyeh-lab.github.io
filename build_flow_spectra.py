#!/usr/bin/env python3
"""Fetch real fluorophore spectra from FPbase and write flow-spectra-data.json.

Run periodically to refresh; the JSON file is committed to the site repo and
loaded by flow.html at startup. Where FPbase has data, flow.html plots the real
spectrum; where it doesn't, it falls back to the existing Gaussian
approximation.

Source: FPbase GraphQL API (https://www.fpbase.org/graphql/).
FPbase data is CC-BY-4.0 — see https://www.fpbase.org/about/.

Usage:
    python3 build_flow_spectra.py
"""

import json
import sys
import urllib.request

ENDPOINT = "https://www.fpbase.org/graphql/"
OUTPUT = "flow-spectra-data.json"

# Map from our flow.html fluorophore name -> FPbase entity
# kind: 'dye' (use Query.dyes by name) or 'protein' (use Query.protein by slug)
# An entry of None means "no FPbase match — keep Gaussian approximation"
MAPPING = {
    # UV (355) — BD Horizon BUV
    "BUV395":             ("dye",     "BD Horizon BUV395"),
    "BUV496":             ("dye",     "BD Horizon BUV496"),
    "BUV563":             ("dye",     "BD Horizon BUV563"),
    "BUV615":             None,
    "BUV661":             ("dye",     "BD Horizon BUV661"),
    "BUV737":             ("dye",     "BD Horizon BUV737"),
    "BUV805":             ("dye",     "BD Horizon BUV805"),
    "DAPI":               ("dye",     "DAPI"),
    "Hoechst 33342":      ("dye",     "Hoechst 33342"),
    "Indo-1 (free)":      None,
    "Indo-1 (bound)":     None,

    # Violet (405) — BV / Pacific / Super Bright / Zombie
    "BV421":              ("dye",     "Brilliant Violet 421"),
    "Pacific Blue":       ("dye",     "Pacific Blue"),
    "eFluor 450":         ("dye",     "eFluor 450"),
    "BV480":              ("dye",     "Brilliant Violet 480"),
    "BV510":              ("dye",     "Brilliant Violet 510"),
    "V500":               ("dye",     "BD Horizon V500"),
    "eFluor 506":         ("dye",     "eFluor 506"),
    "BV570":              ("dye",     "Brilliant Violet 570"),
    "BV605":              ("dye",     "Brilliant Violet 605"),
    "BV650":              ("dye",     "Brilliant Violet 650"),
    "BV711":              ("dye",     "Brilliant Violet 711"),
    "BV750":              ("dye",     "Brilliant Violet 750"),
    "BV785":              ("dye",     "Brilliant Violet 785"),
    "Super Bright 436":   ("dye",     "Super Bright 436"),
    "Super Bright 600":   ("dye",     "Super Bright 600"),
    "Super Bright 645":   ("dye",     "Super Bright 645"),
    "Super Bright 702":   ("dye",     "Super Bright 702"),
    "Super Bright 780":   ("dye",     "Super Bright 780"),
    "LIVE/DEAD Aqua":     ("dye",     "LIVE/DEAD Fixable Aqua"),
    "Zombie UV":          ("dye",     "Zombie UV"),
    "Zombie Violet":      ("dye",     "Zombie Violet"),
    "Zombie Aqua":        ("dye",     "Zombie Aqua"),

    # Blue (488)
    "FITC":               ("dye",     "Fluorescein (FITC)"),
    "Alexa Fluor 488":    ("dye",     "Alexa Fluor 488"),
    "GFP":                ("protein", "EGFP"),
    "eGFP":               ("protein", "EGFP"),
    "mVenus":             ("protein", "mVenus"),
    "YFP":                ("protein", "EYFP"),
    "CFSE":               ("dye",     "CFSE (fluorescein, CFDA, SE)"),
    "PerCP":              ("dye",     "PerCP"),
    "PerCP-Cy5.5":        ("dye",     "PerCP-Cy5.5"),
    "PerCP-eFluor 710":   ("dye",     "PerCP-eFluor 710"),
    "7-AAD":              ("dye",     "7-AAD"),

    # Yellow-Green (561)
    "PE":                 ("dye",     "PE (R-PE / R-phycoerythrin)"),
    "PE-CF594":           ("dye",     "PE-CF594"),
    "PE-Dazzle 594":      ("dye",     "PE/Dazzle 594"),
    "PE-Texas Red":       ("dye",     "PE-Texas Red"),
    "PI":                 ("dye",     "Propidium iodide (PI)"),
    "mCherry":            ("protein", "mCherry"),
    "tdTomato":           ("protein", "tdTomato"),
    "RFP":                ("protein", "mRFP1"),
    "PE-Cy5":             ("dye",     "PE-Cy5"),
    "PE-Cy5.5":           ("dye",     "PE-Cy5.5"),
    "PE-Cy7":             ("dye",     "PE-Cy7"),
    "PE-Fire 640":        None,
    "PE-Fire 700":        None,
    "PE-Fire 744":        None,
    "PE-Fire 780":        ("dye",     "PE/Fire 780"),
    "PE-Fire 810":        None,
    "LIVE/DEAD Yellow":   ("dye",     "LIVE/DEAD Fixable Yellow"),

    # Red (633/640)
    "eFluor 660":         ("dye",     "eFluor 660"),
    "APC":                ("dye",     "APC (allophycocyanin)"),
    "Alexa Fluor 647":    ("dye",     "Alexa Fluor 647"),
    "Alexa Fluor 700":    ("dye",     "Alexa Fluor 700"),
    "APC-R700":           None,
    "APC-Fire 750":       ("dye",     "APC/Fire-750"),
    "APC-Fire 810":       None,
    "APC-H7":             ("dye",     "APC/H7"),
    "APC-Cy7":            ("dye",     "APC/Cy7"),
    "Alexa Fluor 750":    ("dye",     "Alexa Fluor 750"),
    "LIVE/DEAD Far Red":  ("dye",     "LIVE/DEAD Fixable Far Red"),
    "LIVE/DEAD Near-IR":  ("dye",     "LIVE/DEAD Fixable Near-IR"),
}


def gql(query):
    req = urllib.request.Request(
        ENDPOINT,
        data=json.dumps({"query": query}).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "User-Agent": "ayeh-lab-flow-builder/1.0 (https://acyeh-lab.github.io/flow.html)",
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.load(resp)


def downsample(points, step=2):
    """Reduce point density to keep JSON file small. FPbase often has 1nm
    resolution; 2nm is plenty for visual rendering."""
    if not points:
        return points
    out = [points[0]]
    for x, y in points[1:]:
        if x - out[-1][0] >= step:
            out.append([x, y])
    if out[-1][0] != points[-1][0]:
        out.append(points[-1])
    return out


def extract_ex_em(spectra):
    """From a list of FPbase Spectrum objects, return (ex_data, em_data) where
    each is a list of [wavelength, intensity] pairs or None if not present."""
    ex, em = None, None
    for sp in spectra or []:
        sub = (sp.get("subtype") or "").upper()
        # FPbase uses subtypes EX, AB, EM, A_2P, etc.
        # Prefer EX over AB (AB = absorption, EX = excitation; very similar but EX is what we want).
        if sub == "EX" and ex is None:
            ex = sp.get("data")
        elif sub == "AB" and ex is None:
            # Fall back to absorption if no excitation curve
            ex = sp.get("data")
        elif sub == "EM" and em is None:
            em = sp.get("data")
    return downsample(ex) if ex else None, downsample(em) if em else None


def fetch_dye_index():
    """Fetch all dyes once (lookup by name)."""
    print("Fetching all FPbase dyes (this is ~15 MB)...", file=sys.stderr)
    resp = gql("{ dyes { id name slug exMax emMax spectra { subtype data } } }")
    if "errors" in resp:
        raise RuntimeError(f"GraphQL errors: {resp['errors']}")
    return {d["name"]: d for d in resp["data"]["dyes"]}


def fetch_protein(name):
    """Fetch a single protein's default state spectra by exact name."""
    q = (
        '{ allProteins(name_Iexact: "' + name + '") { edges { node '
        "{ name slug defaultState { exMax emMax spectra { subtype data } } } } } }"
    )
    resp = gql(q)
    if "errors" in resp:
        return None
    edges = resp.get("data", {}).get("allProteins", {}).get("edges") or []
    if not edges:
        return None
    node = edges[0]["node"]
    if not node.get("defaultState"):
        return None
    return node


def main():
    out = {
        "_meta": {
            "source": "FPbase GraphQL API (https://www.fpbase.org/graphql/)",
            "license": "CC-BY-4.0",
            "generated_by": "build_flow_spectra.py",
        },
        "fluorophores": {},
    }

    dye_index = fetch_dye_index()
    print(f"Loaded {len(dye_index)} dye states from FPbase.", file=sys.stderr)

    for our_name, mapping in MAPPING.items():
        if mapping is None:
            print(f"  - {our_name}: no FPbase mapping (Gaussian fallback)", file=sys.stderr)
            continue
        kind, key = mapping
        ex, em = None, None
        ex_max, em_max = None, None
        if kind == "dye":
            d = dye_index.get(key)
            if not d:
                print(f"  ! {our_name}: dye {key!r} not found in FPbase", file=sys.stderr)
                continue
            ex, em = extract_ex_em(d.get("spectra"))
            ex_max, em_max = d.get("exMax"), d.get("emMax")
        elif kind == "protein":
            p = fetch_protein(key)
            if not p:
                print(f"  ! {our_name}: protein {key!r} not found in FPbase", file=sys.stderr)
                continue
            st = p["defaultState"]
            ex, em = extract_ex_em(st.get("spectra"))
            ex_max, em_max = st.get("exMax"), st.get("emMax")

        if ex or em:
            out["fluorophores"][our_name] = {
                "fpbase_name": key,
                "kind": kind,
                "ex_max": ex_max,
                "em_max": em_max,
                "ex": ex,
                "em": em,
            }
            ex_pts = len(ex) if ex else 0
            em_pts = len(em) if em else 0
            print(f"  ✓ {our_name}: ex={ex_pts}pts em={em_pts}pts", file=sys.stderr)
        else:
            print(f"  ! {our_name}: matched {key!r} but no usable spectra", file=sys.stderr)

    with open(OUTPUT, "w") as f:
        json.dump(out, f, separators=(",", ":"))
    n = len(out["fluorophores"])
    total = len(MAPPING)
    print(f"\nWrote {OUTPUT}: {n}/{total} fluorophores have real spectra", file=sys.stderr)


if __name__ == "__main__":
    main()
