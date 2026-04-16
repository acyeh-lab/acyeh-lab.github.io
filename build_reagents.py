#!/usr/bin/env python3
"""Parse Hill lab antibody xlsx files into lab-reagents-data.json for flow.html.

Reads:  /fh/fast/hill_g/acyeh_lab/reference_data/hill_reagents/*.xlsx
Writes: lab-reagents-data.json  (committed to the site repo, fetched by flow.html)

Handles the common spreadsheet idiom used by this lab:
- Rows with a "XYZ conjugated antibodies" banner start a fluorochrome section
- A header row "Antibody | Clone | Isotype | Cat Number | Supplier | Notes"
- Data rows follow until the next banner or end of sheet
- UV-conjugate section at the end has the fluorochrome embedded in col 0
  (e.g., "BUV395 anti-CD45") — extracted with regex

stdlib only; no venv needed. Re-run after editing the xlsx files.

Usage:
    python3 build_reagents.py
"""

import json
import re
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

REAGENT_DIR = Path("/fh/fast/hill_g/acyeh_lab/reference_data/hill_reagents")
OUTPUT = Path("lab-reagents-data.json")
NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
NS_T = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"

# Ordered: longer / more-specific patterns come first so they win the first match.
# Each entry maps a section banner (regex) to a canonical fluorochrome name.
# Names that match an entry in flow.html's FLUOROPHORES array let the UI
# add the reagent to the panel with real spectra; unmatched ones are still
# browsable but flagged "no spectrum mapping".
SECTION_PATTERNS = [
    # Red laser specialties
    (r"APC[\s-]?eFluor\s*780",       "APC-Cy7"),           # APC-eFluor 780 ≈ APC-Cy7 emission
    (r"APC[\s/-]?Fire\s*750",        "APC-Fire 750"),
    (r"APC[\s/-]?Fire\s*810",        "APC-Fire 810"),
    (r"APC[\s/-]?Cy7",               "APC-Cy7"),
    (r"APC[\s/-]?H7",                "APC-H7"),
    (r"APC[\s-]?R700",               "APC-R700"),
    (r"eFluor\s*660",                "eFluor 660"),
    (r"Alexa\s*Fluor\s*647",         "Alexa Fluor 647"),
    (r"Alexa\s*Fluor\s*700",         "Alexa Fluor 700"),
    (r"Alexa\s*Fluor\s*750",         "Alexa Fluor 750"),
    (r"\bAPC\b[^/\-]*conjugat",      "APC"),

    # Yellow-green laser
    (r"PE[\s/-]?Cy5\.5",             "PE-Cy5.5"),
    (r"PE[\s/-]?Cy5",                "PE-Cy5"),
    (r"PE[\s/-]?Cy7",                "PE-Cy7"),
    (r"PE[\s/-]?Dazzle\s*594",       "PE-Dazzle 594"),
    (r"PE[\s/-]?CF594",              "PE-CF594"),
    (r"PE[\s/-]?Texas\s*Red",        "PE-Texas Red"),
    (r"PE[\s/-]?eFluor\s*610",       "PE-Texas Red"),     # same channel
    (r"PE[\s/-]?eFluor\s*710",       "PE-Cy5.5"),
    (r"PE[\s/-]?Fire\s*640",         "PE-Fire 640"),
    (r"PE[\s/-]?Fire\s*700",         "PE-Fire 700"),
    (r"PE[\s/-]?Fire\s*744",         "PE-Fire 744"),
    (r"PE[\s/-]?Fire\s*780",         "PE-Fire 780"),
    (r"PE[\s/-]?Fire\s*810",         "PE-Fire 810"),
    (r"\bPE\b[^/\-]*conjugat",       "PE"),

    # Blue laser
    (r"Alexa\s*Fluor\s*488",         "Alexa Fluor 488"),
    (r"\bBB515\b",                   "FITC"),             # BB515 sits in FITC channel
    (r"\bBB700\b",                   "PerCP-Cy5.5"),      # BB700 ≈ PerCP-Cy5.5 channel
    (r"\bFITC\b",                    "FITC"),
    (r"PerCP[\s/-]?Cy5\.5",          "PerCP-Cy5.5"),
    (r"PerCP[\s/-]?eFluor\s*710",    "PerCP-eFluor 710"),
    (r"\bPerCP\b",                   "PerCP"),

    # Violet laser
    (r"Pacific\s*Blue",              "Pacific Blue"),
    (r"eFluor\s*450",                "eFluor 450"),
    (r"eFluor\s*506",                "eFluor 506"),
    (r"\bBV\s*421\b",                "BV421"),
    (r"\bBV\s*480\b",                "BV480"),
    (r"\bBV\s*510\b",                "BV510"),
    (r"\bBV\s*570\b",                "BV570"),
    (r"\bBV\s*605\b",                "BV605"),
    (r"\bBV\s*650\b",                "BV650"),
    (r"\bBV\s*711\b",                "BV711"),
    (r"\bBV\s*750\b",                "BV750"),
    (r"\bBV\s*78[56]\b",             "BV785"),

    # UV laser — banner matches "UV conjugates" as a container; rows below have
    # the fluorochrome embedded in col 0 and are handled specially.
    (r"UV\s*conjugate",              "_UV_EMBEDDED_"),

    # Secondaries / reagents
    (r"Streptavidin",                "_STREP_EMBEDDED_"),  # fluor is in the name
    (r"Annexin\s*V",                 "_ANNEXIN_EMBEDDED_"),
    (r"Biotin\s+conjugat",           "Biotin"),
    (r"biotinylated",                "Biotin"),

    # Tetramers — same idea, each row names its fluorochrome
    (r"PE\s*conjugated\s*tetramer",  "PE"),
    (r"AF488\s*conjugated\s*Tetram", "Alexa Fluor 488"),
    (r"APC\s*conjugated\s*tetramer", "APC"),
    (r"BV\s*421\s*conjugated\s*tetr","BV421"),
    (r"BUV805\s*conjugated\s*tetram","BUV805"),
    (r"APC/Cy7\s*conjugated\s*tetra","APC-Cy7"),
]

# Regex to pull fluorochrome prefix from antibody names like "BUV395 anti-CD45"
EMBEDDED_FLUOR_RE = re.compile(
    r"^\s*(BUV\d{3}|BV\s*\d{3}|PE[\s/-]?Cy\d(?:\.\d)?|PE[\s/-]?Fire\s*\d+|"
    r"APC[\s/-]?Cy\d|APC[\s/-]?H7|APC[\s/-]?Fire\s*\d+|APC[\s-]?R700|"
    r"Alexa\s*Fluor\s*\d+|AF\d+|BB\d+|FITC|PE|APC|"
    r"PerCP[\s/-]?Cy\d\.\d|PerCP|Pacific\s*Blue|eFluor\s*\d+)\s*",
    re.IGNORECASE,
)

FLUOR_NORMALIZE = {
    # Normalize any extracted variant to the canonical FLUOROPHORES name
    "BV421": "BV421", "BV480": "BV480", "BV510": "BV510", "BV570": "BV570",
    "BV605": "BV605", "BV650": "BV650", "BV711": "BV711", "BV750": "BV750",
    "BV785": "BV785", "BV786": "BV785",
    "BUV395": "BUV395", "BUV496": "BUV496", "BUV563": "BUV563",
    "BUV615": "BUV615", "BUV661": "BUV661", "BUV737": "BUV737", "BUV805": "BUV805",
    "PE": "PE", "FITC": "FITC", "APC": "APC", "PerCP": "PerCP",
    "Pacific Blue": "Pacific Blue",
    "AF488": "Alexa Fluor 488", "AF647": "Alexa Fluor 647",
    "AF700": "Alexa Fluor 700", "AF750": "Alexa Fluor 750",
}

# Known bad reagents — from reference_hill_reagents memory. Match by (target, clone, fluor).
KNOWN_BAD = [
    # (fluor, target_substring, clone_substring, reason)
    ("FITC", "BCMA",       "",         "doesn't work (R&D 161616)"),
    ("APC",  "CD44",       "IM7",      "too bright, spills horribly"),
    ("Biotin","CD138",     "281-2",    "too bright per SM"),
    ("FITC", "CD3",        "145-2C11", "doesn't work well"),
    ("APC",  "CD304",      "3E12",     "doesn't work at all (Neuropilin-1)"),
    ("APC",  "Neuropilin", "",         "doesn't work at all (CD304)"),
    ("APC",  "CD45.2",     "104",      "doesn't work well per SM/MK/ZP (BD)"),
]


# --------------- xlsx stdlib parser -----------------------------------------

def col_index(ref):
    letters = re.match(r"([A-Z]+)", ref).group(1)
    n = 0
    for c in letters:
        n = n * 26 + (ord(c) - 64)
    return n - 1


def row_index(ref):
    return int(re.search(r"(\d+)", ref).group(1)) - 1


def read_shared_strings(z):
    try:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    out = []
    for si in root.findall("x:si", NS):
        out.append("".join(t.text or "" for t in si.iter(NS_T)))
    return out


def read_sheet_rows(z, sheet_path, strings):
    root = ET.fromstring(z.read(sheet_path))
    rows = {}
    for row in root.findall(".//x:row", NS):
        cells = {}
        for c in row.findall("x:c", NS):
            ref = c.attrib["r"]
            t = c.attrib.get("t", "n")
            v = c.find("x:v", NS)
            inl = c.find("x:is", NS)
            if v is not None:
                val = strings[int(v.text)] if t == "s" else v.text
            elif inl is not None:
                val = "".join(tt.text or "" for tt in inl.iter(NS_T))
            else:
                continue
            cells[col_index(ref)] = (val or "").strip()
        if cells:
            first_ref = row.find("x:c", NS).attrib["r"]
            rows[row_index(first_ref)] = cells
    return rows


def sheet_paths(z):
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rels_map = {r.attrib["Id"]: r.attrib["Target"] for r in rels}
    out = []
    for s in wb.findall(".//x:sheet", NS):
        rid = s.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
        target = rels_map[rid]
        if not target.startswith("xl/"):
            target = "xl/" + target.lstrip("/")
        out.append((s.attrib["name"], target))
    return out


# --------------- Reagent extraction -----------------------------------------

def match_section(text):
    """Return canonical fluorochrome name for a section-header text, or None."""
    if not text:
        return None
    for pattern, canonical in SECTION_PATTERNS:
        if re.search(pattern, text, re.I):
            return canonical
    return None


def extract_embedded_fluor(name):
    """Pull the fluorochrome prefix out of a name like 'BUV395 anti-CD45'.
    Returns (canonical_fluor, remaining_target) or (None, name)."""
    m = EMBEDDED_FLUOR_RE.match(name or "")
    if not m:
        return None, name
    raw = re.sub(r"\s+", "", m.group(1)).upper()
    # Map variants
    canon = FLUOR_NORMALIZE.get(raw) or FLUOR_NORMALIZE.get(m.group(1).strip()) or m.group(1).strip()
    target = name[m.end():].strip()
    target = re.sub(r"^anti[-\s]*", "", target, flags=re.I)
    target = re.sub(r"^(rat|mouse|hamster|armenian hamster|human)\s+", "", target, flags=re.I)
    return canon, target


def clean_target(s):
    s = (s or "").strip()
    # Normalize whitespace, drop trailing commas/space
    s = re.sub(r"\s+", " ", s)
    return s.rstrip(",;:")


def is_bad(fluor, target, clone):
    target_l = (target or "").lower()
    clone_l = (clone or "").lower()
    for bf, bt, bc, reason in KNOWN_BAD:
        if fluor and bf and bf.lower() not in (fluor or "").lower():
            continue
        if bt and bt.lower() not in target_l:
            continue
        if bc and bc.lower() not in clone_l:
            continue
        return reason
    return None


def parse_file(xlsx_path, species):
    """Extract all reagent rows. species: 'mouse' / 'human' / 'tetramer' / 'isotype'."""
    is_tetramer_file = (species == "tetramer")
    reagents = []
    with zipfile.ZipFile(xlsx_path) as z:
        strings = read_shared_strings(z)
        sheets = sheet_paths(z)
        for sheet_name, sheet_target in sheets:
            # Skip admin sheets
            if sheet_name.lower() in ("ordering checks", "low-use box", "sheet1", "secondary ab"):
                continue
            # Tetramers file has "Mouse" and "Human" sheets — override species
            sp = species
            if is_tetramer_file:
                sp = "mouse tetramer" if "mouse" in sheet_name.lower() else "human tetramer"

            rows = read_sheet_rows(z, sheet_target, strings)
            current_fluor = None
            in_header_area = True
            for r in sorted(rows.keys()):
                cells = rows[r]
                col0 = cells.get(0, "")
                col1 = cells.get(1, "")
                col2 = cells.get(2, "")
                col3 = cells.get(3, "")
                col4 = cells.get(4, "")
                col5 = cells.get(5, "")

                # Skip column-header rows — "Antibody" or "Tetramer" as the
                # target is never a real reagent name. Tetramer sheets use
                # col1="Order ID" instead of "Clone", so don't gate on col1.
                if col0 in ("Antibody", "Tetramer"):
                    in_header_area = False
                    continue

                # Is this a section banner?
                # Heuristic: col0 has a fluorochrome-matching text AND other
                # columns (clone/cat#) are empty. Lets a row like "PE CD45"
                # that IS data still be treated as data.
                looks_like_banner = (
                    not col1 and not col3 and col0
                    and len(col0) < 60
                    and re.search(r"(conjugat|UV conjugates|Streptavidin|Annexin|Biotin|tetramer)", col0, re.I)
                )
                sec = match_section(col0) if looks_like_banner else None
                if sec:
                    current_fluor = sec
                    in_header_area = True
                    continue

                # Commentary / blank filter:
                # - Tetramer sheets: many rows have only col0 (target name) populated,
                #   so we accept as data anything with a col0 that isn't a header.
                # - Antibody sheets: require col1 (clone) or col3 (cat#) so we don't
                #   pull in legend rows like "yellow fill = we have backup".
                if not col0:
                    continue
                if not is_tetramer_file and not col1 and not col3:
                    continue

                # Skip titular rows like "MOUSE ANTIBODIES" banners (caps, long)
                if col0.isupper() and len(col0) > 5 and not col1:
                    continue

                # Determine fluorochrome for this row
                fluor = current_fluor
                target = col0
                if fluor == "_UV_EMBEDDED_" or fluor == "_STREP_EMBEDDED_" or fluor == "_ANNEXIN_EMBEDDED_":
                    embedded, remainder = extract_embedded_fluor(col0)
                    fluor = embedded or fluor
                    target = remainder or col0
                if not fluor:
                    # Try embedded anyway
                    embedded, remainder = extract_embedded_fluor(col0)
                    if embedded:
                        fluor = embedded
                        target = remainder or col0

                if not fluor or fluor.startswith("_"):
                    # Placeholder section without a row-level fluorochrome
                    # embedded in the name (e.g., in-house hybridoma supernatant
                    # rows in the UV section). Skip — not useful in flow panel
                    # design without knowing the conjugate.
                    continue

                # Target-based species override for tetramers: anything with
                # MHC-I (H-2*, H2-*) or MHC-II (I-A(b), I-A(d), I-E*) restriction
                # is a mouse tetramer regardless of which sheet it landed on.
                effective_sp = sp
                if is_tetramer_file:
                    t_upper = target.upper()
                    if (t_upper.startswith("I-A(") or t_upper.startswith("I-AB")
                            or t_upper.startswith("I-E") or t_upper.startswith("H-2")
                            or t_upper.startswith("H2-")):
                        effective_sp = "mouse tetramer"

                # Tetramer sheets put the NIH order ID in col1 and supplier in col2.
                # For tetramers, treat col1 as catalog/order-id, col2 as supplier.
                if is_tetramer_file:
                    clone_val = ""
                    isotype_val = ""
                    catalog_val = col1 or col3
                    supplier_val = col2 or col4 or "NIH"
                    notes_val = col5
                else:
                    clone_val = col1
                    isotype_val = col2
                    catalog_val = col3
                    supplier_val = col4
                    notes_val = col5

                reagent = {
                    "species":     effective_sp,
                    "fluorochrome": fluor,
                    "target":       clean_target(target),
                    "clone":        clone_val,
                    "isotype":      isotype_val,
                    "catalog":      catalog_val,
                    "supplier":     supplier_val,
                    "notes":        notes_val,
                }
                bad = is_bad(fluor, target, col1)
                if bad:
                    reagent["bad"] = bad
                reagents.append(reagent)
    return reagents


def main():
    if not REAGENT_DIR.exists():
        print(f"ERROR: reagent dir not found: {REAGENT_DIR}", file=sys.stderr)
        return 1

    all_reagents = []
    sources = {
        "Mouse antibodies.xlsx":           "mouse",
        "Human antibodies.xlsx":           "human",
        "Mouse and Human tetramers.xlsx":  "tetramer",
        "Isotype Control antibodies.xlsx": "isotype",
    }
    for fname, species in sources.items():
        path = REAGENT_DIR / fname
        if not path.exists():
            print(f"  skip (missing): {fname}", file=sys.stderr)
            continue
        reagents = parse_file(path, species)
        print(f"  {fname}: {len(reagents)} reagents", file=sys.stderr)
        all_reagents.extend(reagents)

    out = {
        "_meta": {
            "source_dir": str(REAGENT_DIR),
            "generated":  time.strftime("%Y-%m-%dT%H:%M:%S"),
            "note":       "Built by build_reagents.py from the Hill lab xlsx inventory files. "
                          "Re-run after editing the xlsx to refresh. Bad-reagent flags come "
                          "from the KNOWN_BAD list in the builder — edit there to add/remove.",
        },
        "reagents": all_reagents,
    }
    with open(OUTPUT, "w") as f:
        json.dump(out, f, separators=(",", ":"))
    print(f"\nWrote {OUTPUT}: {len(all_reagents)} reagents total", file=sys.stderr)
    # Summary by fluor
    by_fluor = {}
    for r in all_reagents:
        by_fluor[r["fluorochrome"]] = by_fluor.get(r["fluorochrome"], 0) + 1
    for fluor, n in sorted(by_fluor.items(), key=lambda x: -x[1])[:15]:
        print(f"  {fluor}: {n}", file=sys.stderr)


if __name__ == "__main__":
    sys.exit(main() or 0)
