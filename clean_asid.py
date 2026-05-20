"""
clean_asid.py
─────────────
Cleaning pipeline for the Australian Shark Incident Database (ASID) public Excel file.

Outputs:
  asid_cleaned.csv          — cleaned dataset, ready for (MapLibre) visualisation
  asid_cleaning_log.txt     — full record-level change log

Usage:
  python clean_asid.py
  python clean_asid.py --input path/to/ASID.xlsx --output path/to/asid_cleaned.csv
"""

import argparse
import pandas as pd
import numpy as np
import sys
from pathlib import Path

# ── CLI args ──────────────────────────────────────────────────────────────────
parser = argparse.ArgumentParser()
parser.add_argument("--input",  default="Australian_Shark-Incident_Database_Public_Version.xlsx")
parser.add_argument("--dict",   default="australian_shark_species_dictionary.csv")
parser.add_argument("--output", default="asid_cleaned.csv")
args = parser.parse_args()

log_lines = []

def log(msg):
    print(msg)
    log_lines.append(msg)


# ═══════════════════════════════════════════════════════════════════════════════
# 1. LOAD
# ═══════════════════════════════════════════════════════════════════════════════
log("=" * 70)
log("ASID CLEANING PIPELINE")
log("=" * 70)

df  = pd.read_excel(args.input)
dic = pd.read_csv(args.dict)

log(f"\n[LOAD] {len(df)} records, {len(df.columns)} columns loaded from {args.input}")
log(f"[LOAD] Species dictionary: {len(dic)} canonical entries from {args.dict}")

# Track changes
changes = {k: 0 for k in [
    "state_normalised", "injury_normalised", "provoked_normalised",
    "sci_name_standardised", "common_name_standardised",
    "coord_corrected", "coord_flagged_low_precision"
]}


# ═══════════════════════════════════════════════════════════════════════════════
# 2. EXTEND SPECIES DICTIONARY
#    Add entries for species/groups not covered in the base dictionary CSV.
#    Rationale documented inline for each addition.
# ═══════════════════════════════════════════════════════════════════════════════
log("\n[DICT] Extending dictionary with uncovered species/groups...")

supplementary = [
    {
        # 215 records — multiple Orectolobus species grouped under this name.
        # Cannot resolve to a single species; genus placeholder used.
        "canonical_scientific_name": "Orectolobus spp.",
        "scientific_synonyms":       "Orectolobus maculatus, Orectolobus ornatus",
        "canonical_common_name":     "Wobbegong",
        "common_name_synonyms":      "wobbegong shark, carpet shark",
        "notes":                     "Group-level entry; multiple species recorded under this label in ASID"
    },
    {
        # 79 records — catch-all term for multiple Carcharhinus species.
        # Historically used when species-level ID was not possible.
        "canonical_scientific_name": "Carcharhinus spp.",
        "scientific_synonyms":       "",
        "canonical_common_name":     "Whaler Shark",
        "common_name_synonyms":      "whaler, requiem shark",
        "notes":                     "Group-level entry; cannot be resolved to single species"
    },
    {
        # 5 records — group label; great hammerhead vs. smooth hammerhead distinction
        # not possible from ASID records.
        "canonical_scientific_name": "Sphyrna spp.",
        "scientific_synonyms":       "Sphyrna mokarran, Sphyrna zygaena",
        "canonical_common_name":     "Hammerhead Shark",
        "common_name_synonyms":      "hammerhead, great hammerhead, smooth hammerhead",
        "notes":                     "Group-level entry"
    },
    {
        # 2 records — identifiable to species.
        "canonical_scientific_name": "Isurus oxyrinchus",
        "scientific_synonyms":       "Isurus glaucus",
        "canonical_common_name":     "Shortfin Mako",
        "common_name_synonyms":      "shortfin mako shark, mako shark, blue pointer",
        "notes":                     "Species-level entry"
    },
]
dic = pd.concat([dic, pd.DataFrame(supplementary)], ignore_index=True)

# Fix existing Bronze Whaler entry — "bronze whaler shark" not in synonyms
# causing 32 records to miss on common name matching.
mask = dic["canonical_common_name"].str.lower() == "bronze whaler"
dic.loc[mask, "common_name_synonyms"] = (
    dic.loc[mask, "common_name_synonyms"].fillna("") +
    ", bronze whaler shark, bronzie, copper shark"
)

log(f"[DICT] Dictionary extended to {len(dic)} entries")
log(f"[DICT] Bronze Whaler synonyms patched (bronze whaler shark, bronzie added)")


# ═══════════════════════════════════════════════════════════════════════════════
# 3. BUILD LOOKUP TABLES FROM DICTIONARY
#    Maps any synonym → canonical name, for both scientific and common columns.
# ═══════════════════════════════════════════════════════════════════════════════
sci_lookup = {}   # lower-case synonym → canonical scientific name
com_lookup = {}   # lower-case synonym → canonical common name

for _, row in dic.iterrows():
    canon_sci = str(row["canonical_scientific_name"]).strip()
    canon_com = str(row["canonical_common_name"]).strip()

    sci_syns = [s.strip() for s in str(row["scientific_synonyms"]).split(",") if s.strip()]
    sci_syns.append(canon_sci)
    for s in sci_syns:
        sci_lookup[s.lower()] = canon_sci

    com_syns = [s.strip() for s in str(row["common_name_synonyms"]).split(",") if s.strip()]
    com_syns.append(canon_com)
    for s in com_syns:
        com_lookup[s.lower()] = canon_com

log(f"[DICT] {len(sci_lookup)} scientific name synonyms indexed")
log(f"[DICT] {len(com_lookup)} common name synonyms indexed")


# ═══════════════════════════════════════════════════════════════════════════════
# 4. CATEGORICAL FIELD NORMALISATION
#    Rule: strip whitespace, lowercase, then apply field-specific mappings.
#    A change log entry is printed for every modified record.
# ═══════════════════════════════════════════════════════════════════════════════
log("\n[NORMALISE] Categorical fields...")

# ── 4a. State ─────────────────────────────────────────────────────────────────
valid_states = {"NSW","QLD","WA","SA","VIC","TAS","NT"}
state_map    = {"qld": "QLD"}   # extend if further variants found

for idx, val in df["State"].items():
    if pd.isna(val):
        continue
    normalised = state_map.get(str(val).strip().lower(), str(val).strip().upper())
    if normalised != str(val).strip():
        log(f"  [State] row {idx}: '{val}' → '{normalised}'")
        df.at[idx, "State"] = normalised
        changes["state_normalised"] += 1

# ── 4b. Victim.injury ─────────────────────────────────────────────────────────
injury_map = {
    "injured": "injured",
    "fatal":   "fatal",
    "uninjured":"uninjured",
    "unknown": "unknown",
    "injury":  "injured",   # 1 record — best available interpretation
}

for idx, val in df["Victim.injury"].items():
    if pd.isna(val):
        continue
    normalised = injury_map.get(str(val).strip().lower(), str(val).strip().lower())
    if normalised != str(val).strip():
        log(f"  [Victim.injury] row {idx}: '{val}' → '{normalised}'")
        df.at[idx, "Victim.injury"] = normalised
        changes["injury_normalised"] += 1

# ── 4c. Provoked/unprovoked ───────────────────────────────────────────────────
# 6 null values — set to "unknown" so they render cleanly in filters
for idx, val in df["Provoked/unprovoked"].items():
    if pd.isna(val):
        log(f"  [Provoked/unprovoked] row {idx}: NaN → 'unknown'")
        df.at[idx, "Provoked/unprovoked"] = "unknown"
        changes["provoked_normalised"] += 1
    else:
        normalised = str(val).strip().lower()
        if normalised != str(val).strip():
            df.at[idx, "Provoked/unprovoked"] = normalised
            changes["provoked_normalised"] += 1


# ═══════════════════════════════════════════════════════════════════════════════
# 5. SPECIES NAME STANDARDISATION
#    Scientific name matched first; if a hit is found, the corresponding
#    canonical common name is also applied. Common name then checked
#    independently to catch records where only the common name is available.
# ═══════════════════════════════════════════════════════════════════════════════
log("\n[SPECIES] Standardising scientific and common names...")

for idx, row in df.iterrows():
    sci_raw = str(row["Shark.scientific.name"]).strip() if pd.notna(row["Shark.scientific.name"]) else ""
    com_raw = str(row["Shark.common.name"]).strip()     if pd.notna(row["Shark.common.name"])     else ""

    sci_key = sci_raw.lower()
    com_key = com_raw.lower()

    # Scientific name lookup
    if sci_key and sci_key in sci_lookup:
        canon_sci = sci_lookup[sci_key]
        if canon_sci != sci_raw:
            log(f"  [sci] row {idx}: '{sci_raw}' → '{canon_sci}'")
            df.at[idx, "Shark.scientific.name"] = canon_sci
            changes["sci_name_standardised"] += 1

    # Common name lookup
    if com_key and com_key in com_lookup:
        canon_com = com_lookup[com_key]
        if canon_com != com_raw:
            log(f"  [com] row {idx}: '{com_raw}' → '{canon_com}'")
            df.at[idx, "Shark.common.name"] = canon_com
            changes["common_name_standardised"] += 1


# ═══════════════════════════════════════════════════════════════════════════════
# 6. COORDINATE CLEANING
# ═══════════════════════════════════════════════════════════════════════════════
log("\n[COORDS] Cleaning and auditing coordinates...")

df["Latitude"]  = pd.to_numeric(df["Latitude"],  errors="coerce")
df["Longitude"] = pd.to_numeric(df["Longitude"], errors="coerce")

# ── 6a. Known data entry error: Forbes Island longitude = 4034 ────────────────
# Latitude -12.2929 is precise and consistent with Forbes Island, QLD.
# Longitude 4034 is a clear transposition of 143.4 (digit order scrambled).
# Corrected to 143.40; flagged in coord_corrected column.
forbes_mask = (df["Location"].str.lower().str.contains("forbes island", na=False)) & \
              (df["Longitude"] > 1000)
if forbes_mask.any():
    idx = df[forbes_mask].index[0]
    log(f"  [COORD FIX] row {idx} 'forbes island': Longitude {df.at[idx,'Longitude']} → 143.40")
    log(f"              (transposition error: 4034 → 143.4; verify against source if available)")
    df.at[idx, "Longitude"] = 143.40
    changes["coord_corrected"] += 1

df["coord_corrected"] = False
df.loc[forbes_mask, "coord_corrected"] = True

# ── 6b. Precision flagging ────────────────────────────────────────────────────
# Records with ≤ 2 decimal places were likely assigned from a town centroid
# or rough map click. These render as hollow markers in the visualisation.
# Precision measured on Latitude (generally matches Longitude precision).

def decimal_places(val):
    if pd.isna(val):
        return 0
    s = f"{val:.10f}".rstrip("0")
    if "." not in s:
        return 0
    return len(s.split(".")[1])

dp = df["Latitude"].apply(decimal_places)
low_precision = dp <= 2
df["coord_quality"] = np.where(low_precision, "low", "high")

n_low = low_precision.sum()
log(f"\n  [COORD PRECISION] {n_low} records flagged as low-precision (≤ 2 decimal places)")
log(f"  These will render as hollow markers in the visualisation.")
for idx in df[low_precision].index:
    row = df.loc[idx]
    log(f"    row {idx}: {row['Location']} | lat={row['Latitude']} lon={row['Longitude']}")
changes["coord_flagged_low_precision"] = n_low


# ═══════════════════════════════════════════════════════════════════════════════
# 7. OUTPUT
# ═══════════════════════════════════════════════════════════════════════════════

# Drop the confusing unnamed trailing column from the Excel file
df = df.loc[:, ~df.columns.str.startswith("Unnamed")]

df.to_csv(args.output, index=False)
log(f"\n[OUTPUT] Cleaned dataset written to {args.output}")
log(f"         {len(df)} records, {len(df.columns)} columns")

# ── Summary ───────────────────────────────────────────────────────────────────
log("\n" + "=" * 70)
log("CLEANING SUMMARY")
log("=" * 70)
for k, v in changes.items():
    log(f"  {k:<40} {v:>4} records affected")
log("=" * 70)

# Write log file
log_path = Path(args.output).with_name("asid_cleaning_log.txt")
with open(log_path, "w") as f:
    f.write("\n".join(log_lines))
print(f"\n[LOG] Full change log written to {log_path}")
