"""
name: Honda DEI export
description: Builds the Honda diversity & fees export for AHM.
required_files:
  BASE_FILE: "base_table.csv"        # billing export
  DEMO_FILE: "demographics.csv"      # HR / demographics export
outputs:
  OUTPUT_FILE: "honda_export.xlsx"   # final deliverable
"""

"""
Build the Honda‑AHM diversity & fees export.

Inputs
------
1. base_table.csv   – Billing rows (Person Code, Person Name, Personnel Type Type, Office, Matter Code, PF/PR metrics …)
2. demographics.csv – HR attributes keyed by Aderant Number (gender, ethnicity, orientation, consent flag …)

Output
------
honda_export.xlsx – Columns:
  • Attorney Name (Firm)
  • Honda Attorney Name (retained by)
  • AHM Legal Assigned Matter Number
  • Total Fees (no experts, courts reporters etc.)
  • Gender (drop down)
  • Firm Position (drop down)
  • Diversity Profile
  • Additional Diversity Profile Information
"""

import sys
import os            
from pathlib import Path
import re
import pandas as pd

# Folder where input/output files live
# • When frozen, the EXE runs from a temp dir, so use the folder you launched it from (cwd)
# • When running the .py file directly, use the script’s own folder
if getattr(sys, "frozen", False):          # PyInstaller EXE
    SCRIPT_DIR = Path(os.getcwd()).resolve()
else:                                      # .py script
    SCRIPT_DIR = Path(__file__).resolve().parent

# ───────────────────────────────────────────────────────────
# CONFIG – adjust for your environment
# ───────────────────────────────────────────────────────────
CONFIG = {
    "BASE_FILE": "base_table.csv",         # billing export
    "DEMO_FILE": "demographics.csv",       # HR/demographics export
    "OUTPUT_FILE": "honda_export.xlsx",     # final deliverable

    # Column headers exactly as they appear in the input files  ───┐
    "COLS": {                                                    # │
        "person_code": "Person Code",                          # │
        "person_name": "Person Name",                          # │
        "personnel_type": "Personnel Type Type",               # │
        "matter_code": "Matter Code",                          # │
        "total_fees": "PF+PR Dollars Billed",                  # │
        # demographics file
        "aderant": "Aderant Number",                           # │ parent key
        "gender_code": "Gender Code (Legal)",                  # │
        "gender_identity": "Gender Identity",                  # │
        "ethnicity": "Ethnicity Code Description",             # │
        "orientation": "Sexual Orientation",                   # │
        "consent": "Consent to Share Demographics Description",# │
    },                                                            # ┘
}

FINAL_COLUMNS = [
    "Attorney Name (Firm)",
    "Honda Attorney Name (retained by)",
    "AHM Legal Assigned Matter Number",
    "Total Fees (no experts, courts reporters etc.) ",
    "Gender (drop down)",
    "Firm Position (drop down)",
    "Diversity Profile ",
    "Additional Diversity Profile Information (drop down)",
]

# ───────────────────────────────────────────────────────────
# Helper look‑ups
# ───────────────────────────────────────────────────────────

_ethnicity_map = {
    "american indian or alaska native": "American Indian/Alaska Native",
    "native american": "American Indian/Alaska Native",
    "african american": "African American",
    "black or african american": "African American",
    "asian": "Asian",
    "hispanic or latino": "Hispanic/Latino",
    "latino": "Hispanic/Latino",
    "multiracial": "Multi-Racial",
    "two or more races": "Multi-Racial",
    "native hawaiian or other pacific islander": "Hawaiian/Other Pacific Islander",
    "hawaiian or other pacific islander": "Hawaiian/Other Pacific Islander",
    "white": "White",
    "white/caucasian": "White",
    "two or more nationality": "Multi-Racial",
    "middle eastern / north african": "",   # no mapping
    "not defined": "",                      # no mapping
}

_orientation_map = {
    "lgbt": "LGBTQ",
    "lgbtq": "LGBTQ",
    "lgbtq+": "LGBTQ",
    "disabled": "Disabled",
}


# ───────────────────────────────────────────────────────────
# Core functions
# ───────────────────────────────────────────────────────────

def read_table(path: Path) -> pd.DataFrame:
    """Load CSV or Excel based on extension, return DataFrame."""
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path)
    if path.suffix.lower() in (".xls", ".xlsx"):
        return pd.read_excel(path, engine="openpyxl")
    raise ValueError(f"Unsupported file type: {path}")


def forward_fill_columns(df: pd.DataFrame, cols):
    """Forward‑fill the ragged label columns in‑place."""
    df[cols] = df[cols].ffill()
    return df


def map_firm_position(raw: str) -> str:
    if pd.isna(raw):
        return "Other"
    val = str(raw).strip().lower()
    if val == "equity partner":
        return "Equity Partner"
    if "partner" in val:
        return "Non Equity Partner"
    if "counsel" in val:
        return "Counsel"
    if "associate" in val:
        return "Associate"
    return "Other"


def extract_matter_suffix(code: str) -> str:
    """Return 4‑digit suffix after dot, preserving leading zeros."""
    if pd.isna(code):
        return ""
    m = re.search(r"\.([0-9]{4})", str(code))
    return m.group(1) if m else ""


def map_gender(row) -> str:
    if row[CONFIG["COLS"]["consent"]] != "Y":
        return ""

    g_code = str(row.get(CONFIG["COLS"]["gender_code"], "")).strip().lower()
    g_id   = str(row.get(CONFIG["COLS"]["gender_identity"], "")).strip().lower()

    if g_code.startswith("m"):
        return "M"
    if g_code.startswith("f"):
        return "F"

    # Treat Non‑Binary or any value containing 'trans' as Transgender
    if "non-binary" in g_id or "trans" in g_id:
        return "Transgender"

    return ""   # Prefer not to say → blank


def map_ethnicity(row) -> str:
    if row[CONFIG["COLS"]["consent"]] != "Y":
        return ""
    val = str(row.get(CONFIG["COLS"]["ethnicity"], "")).strip().lower()

    # Try exact match first
    if val in _ethnicity_map:
        return _ethnicity_map[val]

    # Fuzzy: return the first mapping whose key is contained in the value
    for key, mapped in _ethnicity_map.items():
        if key and key in val:
            return mapped
    return ""


def map_orientation(row) -> str:
    """
    Collapse declared orientations into Honda's single‑choice list.
    Return 'LGBTQ' for any non‑straight orientation when consent = 'Y'.
    """
    if row[CONFIG["COLS"]["consent"]] != "Y":
        return ""

    val = str(row.get(CONFIG["COLS"]["orientation"], "")).strip().lower()

    if val in {
        "", "heterosexual/straight", "prefer not to say",
        "sexual orientation not selected",
    }:
        return ""

    if val == "disabled":
        return "Disabled"

    # Gay/Lesbian, Bisexual, Queer, Asexual, Pansexual, etc.
    return "LGBTQ"


def build_export(base_df: pd.DataFrame, demo_df: pd.DataFrame) -> pd.DataFrame:
    c = CONFIG["COLS"]

    # Forward‑fill ragged columns
    base_df = forward_fill_columns(base_df, [c["person_code"], c["person_name"], c["personnel_type"], "Office"])

    # Join on Person Code / Aderant Number
    merged = base_df.merge(demo_df, how="left", left_on=c["person_code"], right_on=c["aderant"], suffixes=("", "_demo"))

    # Compute each final column
    final = pd.DataFrame()
    final["Attorney Name (Firm)"] = merged[c["person_name"]]
    final["Honda Attorney Name (retained by)"] = ""
    final["AHM Legal Assigned Matter Number"] = merged[c["matter_code"]].apply(extract_matter_suffix)
    final["Total Fees (no experts, courts reporters etc.) "] = merged[c["total_fees"]]
    final["Gender (drop down)"] = merged.apply(map_gender, axis=1)
    final["Firm Position (drop down)"] = merged[c["personnel_type"]].apply(map_firm_position)
    final["Diversity Profile "] = merged.apply(map_ethnicity, axis=1)
    final["Additional Diversity Profile Information (drop down)"] = merged.apply(map_orientation, axis=1)

    # ── Post‑processing ──────────────────────────────────────────
    # 1) Convert fees to numeric for filtering / formatting
    fees_num = pd.to_numeric(
        final["Total Fees (no experts, courts reporters etc.) "], errors="coerce"
    ).fillna(0)

    # 2) Build filters
    mask_fees  = fees_num > 0                                  # keep > 0
    mask_real  = ~final["Attorney Name (Firm)"].str.strip().str.lower().eq(
        "flat fee billing allocation"
    )

    # 3) Apply filters
    keep_mask = mask_fees & mask_real
    final     = final[keep_mask].copy()
    fees_num  = fees_num[keep_mask]

    # 4) Format money exact to the penny
    def _fmt(n):
        return f"${n:,.2f}"
    final["Total Fees (no experts, courts reporters etc.) "] = fees_num.apply(_fmt)

    # 5) Sort for readability
    final.sort_values(
        ["Attorney Name (Firm)", "AHM Legal Assigned Matter Number"],
        inplace=True,
    )

    return final[FINAL_COLUMNS]


def write_table(df: pd.DataFrame, path: Path):
    if path.suffix.lower() == ".csv":
        df.to_csv(path, index=False)
    else:
        df.to_excel(path, index=False, engine="openpyxl")


# ───────────────────────────────────────────────────────────
# Entry‑point
# ───────────────────────────────────────────────────────────

def main(BASE_FILE: Path, DEMO_FILE: Path) -> Path:
    """
    Entry point used by the Qt desktop GUI.

    Parameters
    ----------
    BASE_FILE : Path
        Billing export (CSV/XLSX).
    DEMO_FILE : Path
        Demographics export (CSV/XLSX).

    Returns
    -------
    Path
        Location of the generated Honda export file.
    """
    base_df = read_table(BASE_FILE)
    demo_df = read_table(DEMO_FILE)

    final_df = build_export(base_df, demo_df)

    out_path = BASE_FILE.with_name(CONFIG["OUTPUT_FILE"])
    write_table(final_df, out_path)
    print(f"Export complete → {out_path}")
    return out_path


def cli_main():
    base_df = read_table(SCRIPT_DIR / CONFIG["BASE_FILE"])
    demo_df = read_table(SCRIPT_DIR / CONFIG["DEMO_FILE"])

    final_df = build_export(base_df, demo_df)
    write_table(final_df, SCRIPT_DIR / CONFIG["OUTPUT_FILE"])
    print("Export complete →", CONFIG["OUTPUT_FILE"])


if __name__ == "__main__":
    cli_main()
