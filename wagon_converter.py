"""Consolidate multi-sheet rake/wagon workbooks into the standard 24-column CSV.

One output row per physical item (plate / coil / bar). The richest detail sheet
in the workbook is the priority source; the 11-digit rail wagon number is either
read from that sheet, joined from another sheet by item id (HR coil), or looked
up against the load-plan sheet by the last 5 digits (Plate). No fuzzy matching:
a value is resolved deterministically or left blank, never guessed.
"""

from __future__ import annotations

import re
from collections.abc import Callable
from pathlib import Path

import pandas as pd

Logger = Callable[[str], None]

SCHEMA: list[str] = [
    "S.No.", "Lot No.", "Wagon No.", "Sold To Party.", "Ship To Party.",
    "Ship To Party City.", "SO No.", "Line Item.", "OBD.", "Location.",
    "Material.", "Plant code.", "Heat no.", "Batch Id.", "Material Code.",
    "Grade.", "Size.", "Diameter.", "Thickness.", "Width.",
    "Theoretical. Wt.", "Length.", "Act.ual Wt.", "Pcs.",
]

# Destination column -> source header spellings (normalized), best first.
# Distinct destinations stay distinct: Sold-To vs Ship-To, Theoretical vs Actual.
FIELD_SYNONYMS: dict[str, list[str]] = {
    "Lot No.": ["lotno", "lot"],
    "Sold To Party.": ["soldtoparty", "soldto"],
    "Ship To Party.": ["shiptoparty", "party", "partyname", "customer"],
    "Ship To Party City.": ["shipcity", "shiptocity", "cityshiptoparty", "destination", "city"],
    "SO No.": ["sono", "sostono", "so","salesorder"],
    "Line Item.": ["lineitem", "item", "li", "soitem"],
    "OBD.": ["obdno", "obd"],
    "Location.": ["storagelocation", "sloc", "location", "loc", "yard"],
    "Material.": ["product", "ptype", "material"],
    "Plant code.": ["plantcode", "plant", "payercode"],
    "Heat no.": ["heatno", "heat"],
    "Batch Id.": ["plateid", "coilid", "coilsid", "tmtbarid", "barid", "batchid", "batch"],
    "Material Code.": ["materialcode"],
    "Grade.": ["externalgrade", "grade", "internalgrade"],
    "Size.": ["coilsize", "size"],
    "Diameter.": ["dia", "diameter"],
    "Thickness.": ["thick", "thk", "thickness"],
    "Width.": ["width", "wth"],
    "Theoretical. Wt.": ["netwt", "theoreticalwt", "sapweight","productionqty"],
    "Length.": ["length"],
    "Act.ual Wt.": ["pgiwt", "actualwt", "weight", "wt", "sumofwt","productionqty"],
    "Pcs.": ["pcs", "pieces", "noofcoils", "bundle", "totalqty"],
}

# A real item row has its own id (plate/coil/bar). Footer & group-trailer rows
# lack it; the id is never group-filled, so it stays a reliable discriminator.
_ITEM_KEY = "Batch Id."

# Fields shared by every item in a wagon group (printed once, on the group's
# first or last row). Filled across the group; see _group_fill for direction.
_GROUP_FIELDS = [
    "Wagon No.", "Lot No.", "Sold To Party.", "Ship To Party.",
    "Ship To Party City.", "SO No.", "Line Item.", "OBD.",
]


def normalize(text: object) -> str:
    """Lowercase a header/value to bare alphanumerics for matching."""
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def clean_cell(value: object) -> str:
    """Render a cell as a tidy string: drop float `.0` on integers, expand
    scientific notation, blank out NaN. Keeps real decimals (weights) intact."""
    if pd.isna(value):
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (int, float)):
        return str(int(value)) if float(value).is_integer() else f"{value:g}"
    text = str(value).strip()
    if "e+" in text.lower():
        try:
            return f"{float(text):.0f}"
        except ValueError:
            pass
    return "" if text.lower() == "nan" else text


def extract_wagon(value: object) -> str:
    """Return the 11-digit rail wagon number embedded in a cell, else ''.
    Handles a leading party prefix (e.g. 'JSPL26452614091' -> '26452614091')."""
    digits = re.sub(r"\D", "", clean_cell(value))
    return digits if len(digits) == 11 else ""


def classify_material(row: dict[str, str]) -> str:
    """Material from dimensions:
        diameter present                  -> TMT
        thickness + width + length        -> Plate
        thickness + width                 -> COIL  (CR and HR coil alike)
    Falls back to whatever the source 'Material.' column carried."""
    dia = row.get("Diameter.", "").strip()
    thk = row.get("Thickness.", "").strip()
    wid = row.get("Width.", "").strip()
    lng = row.get("Length.", "").strip()
    if dia:
        return "TMT"
    if thk and wid and lng:
        return "Plate"
    if thk and wid:
        return "COIL"
    return row.get("Material.", "")


def find_header_row(raw: pd.DataFrame) -> int:
    """Pick the row (within the first 15) that maps the most schema fields."""
    known = {syn for syns in FIELD_SYNONYMS.values() for syn in syns}
    known |= {"wagon", "wagonno", "vehicleno", "slno", "sno"}
    best_row, best_score = 0, 0
    for i in range(min(15, len(raw))):
        score = sum(normalize(v) in known for v in raw.iloc[i].to_numpy())
        if score > best_score:
            best_row, best_score = i, score
    return best_row if best_score >= 3 else 0


def load_sheets(path: str, logger: Logger) -> dict[str, pd.DataFrame]:
    """Read every sheet, snapping each to its detected header row."""
    logger(f"Reading {path}")
    book = pd.ExcelFile(path, engine="openpyxl")
    sheets: dict[str, pd.DataFrame] = {}
    for name in book.sheet_names:
        raw = book.parse(name, header=None)
        if raw.empty:
            continue
        df = book.parse(name, skiprows=find_header_row(raw)).dropna(how="all")
        if not df.empty:
            sheets[name] = df
    return sheets


def build_colmap(columns: pd.Index) -> dict[str, object]:
    """Map source columns -> schema destination, one source per destination
    (first synonym wins). Returns {destination: source_column}."""
    norm_cols = {col: normalize(col) for col in columns}
    mapping: dict[str, object] = {}
    for dest, synonyms in FIELD_SYNONYMS.items():
        for syn in synonyms:
            hit = next((c for c, n in norm_cols.items() if n == syn), None)
            if hit is not None:
                mapping[dest] = hit
                break
    return mapping


def score_sheet(df: pd.DataFrame) -> int:
    """How many schema fields a sheet can fill (used to pick the detail sheet)."""
    return len(build_colmap(df.columns))


def wagon_column(df: pd.DataFrame) -> object | None:
    """The column most likely to hold a wagon number, if any. A 'wagon' column
    wins over 'vehicle no' (which is often the road-truck number, not the rake)."""
    for col in df.columns:
        if "wagon" in normalize(col):
            return col
    for col in df.columns:
        if normalize(col) == "vehicleno":
            return col
    return None


def build_wagon_lookup(sheets: dict[str, pd.DataFrame]) -> dict[str, str]:
    """item-id (digits) -> 11-digit wagon, for sheets that pair the two
    (HR coil: 'Final Annexure' links Wagon No. <-> Coil's ID)."""
    lookup: dict[str, str] = {}
    for df in sheets.values():
        wcol = wagon_column(df)
        colmap = build_colmap(df.columns)
        id_col = colmap.get("Batch Id.")
        if wcol is None or id_col is None:
            continue
        wagons = df[wcol].ffill()  # wagon printed once per merged group
        for raw_id, raw_wagon in zip(df[id_col], wagons, strict=True):
            wagon = extract_wagon(raw_wagon)
            key = re.sub(r"\D", "", clean_cell(raw_id))
            if wagon and key:
                lookup.setdefault(key, wagon)
    return lookup


# Brake-van wagon-type codes (Indian Railways): BVZC / BVZI / BVCM — all 'BV…'.
# Brake vans carry the guard, no cargo, so they are dropped from the output.
_BRAKE_VAN_CODES = {"bvzc", "bvzi", "bvcm"}


def _is_brake_van(cell: object) -> bool:
    """True if a cell is a brake-van wagon-type code (BVZC/BVZI/BVCM or any
    short 'BV…' variant). Length-guarded so prose never trips it."""
    n = normalize(cell)
    return n in _BRAKE_VAN_CODES or (n.startswith("bv") and len(n) <= 8)


def brake_van_wagons(sheets: dict[str, pd.DataFrame]) -> set[str]:
    """11-digit wagon numbers whose row carries a brake-van type code, in any
    sheet (load plan, annexure, trailer list — wherever the type is printed)."""
    vans: set[str] = set()
    for df in sheets.values():
        for _, row in df.iterrows():
            cells = [clean_cell(v) for v in row.to_numpy()]
            if not any(_is_brake_van(c) for c in cells):
                continue
            for c in cells:
                digits = re.sub(r"\D", "", c)
                if len(digits) >= 11:
                    vans.add(digits[-11:])
    return vans


def build_loadplan_lookup(sheets: dict[str, pd.DataFrame]) -> dict[str, str]:
    """Plate: the load-plan sheet carries the full 11-digit wagons; the detail
    rows carry only a short loading code whose digits are the wagon's last 5.
    Index each full wagon by its last 5 digits, so a detail row's code resolves
    to the full number (a VLOOKUP on the last 5 -> prepend the first 6).
    Brake-van rows (BV…) are skipped — they hold no cargo.

    ASSUMPTION A: the load-plan sheet name contains 'load'.
    ASSUMPTION B: the full wagon is the rightmost 11 digits of its cell."""
    lookup: dict[str, str] = {}
    for name, df in sheets.items():
        if "load" not in normalize(name):
            continue
        for _, row in df.iterrows():
            cells = [clean_cell(v) for v in row.to_numpy()]
            if any(_is_brake_van(c) for c in cells):
                continue  # skip the brake van's wagon
            for c in cells:
                digits = re.sub(r"\D", "", c)
                if len(digits) >= 11:
                    wagon = digits[-11:]                  # rightmost 11 (drop prefix)
                    lookup.setdefault(wagon[-5:], wagon)  # last 5 -> full 11
    return lookup


def pick_detail_sheet(sheets: dict[str, pd.DataFrame]) -> str:
    """The sheet that maps the most schema fields (ties -> most rows)."""
    return max(sheets, key=lambda n: (score_sheet(sheets[n]), len(sheets[n])))


def _group_fill(detail: pd.DataFrame, colmap: dict[str, object], wcol: object | None) -> None:
    """Spread group-shared fields (wagon, SO, party...) across each wagon group,
    in place. Direction is inferred from where the wagon prints:
      - on item rows (merged top, CR/PLATE/HR) -> forward fill
      - on a separate trailer row (TMT)        -> back fill
    ponytail: a single ffill/bfill per column; assumes groups are contiguous and
    the wagon marks the group edge. Re-check if a workbook interleaves groups."""
    if wcol is None:
        return
    has_wagon = detail[wcol].notna()
    if not has_wagon.any():
        return
    id_col = colmap.get("Batch Id.")
    on_item_rows = id_col is not None and detail.loc[has_wagon, id_col].notna().any()
    fill = pd.DataFrame.ffill if on_item_rows else pd.DataFrame.bfill
    cols = [wcol] + [colmap[d] for d in _GROUP_FIELDS if d in colmap]
    for col in dict.fromkeys(cols):  # de-dup, preserve order
        detail[col] = fill(detail[col])


def run_conversion(
    input_path: str,
    output_schema: list[str],
    output_path: str,
    primary_sheet_name: str | None = None,
    logger: Logger = print,
) -> dict[str, object]:
    """Read the workbook, emit the standard CSV. Signature kept for the GUI."""
    try:
        sheets = load_sheets(input_path, logger)
    except Exception as exc:  # noqa: BLE001 - surfaced to the GUI as a message
        logger(f"Error reading workbook: {exc}")
        return {"error": str(exc)}
    if not sheets:
        return {"error": "No readable sheets found."}

    detail_name = primary_sheet_name if primary_sheet_name in sheets else pick_detail_sheet(sheets)
    detail = sheets[detail_name].copy()
    colmap = build_colmap(detail.columns)
    logger(f"Detail sheet: '{detail_name}' ({len(detail)} rows, {len(colmap)} fields mapped)")

    wcol = wagon_column(detail)
    _group_fill(detail, colmap, wcol)
    id_lookup = build_wagon_lookup(sheets) if wcol is None else {}
    loadplan = build_loadplan_lookup(sheets)
    vans = brake_van_wagons(sheets)
    if id_lookup:
        logger(f"Wagon resolved via item-id link ({len(id_lookup)} ids).")
    if loadplan:
        logger(f"Load plan indexed: {len(loadplan)} wagons by last-5 digits.")
    if vans:
        logger(f"Brake vans found and excluded: {len(vans)}.")

    rows: list[dict[str, str]] = []
    failed = 0
    # ponytail: iterrows is fine at these sizes (<1k rows); vectorize if files grow.
    for _, row in detail.iterrows():
        out = {col: "" for col in output_schema}
        for dest, src in colmap.items():
            out[dest] = clean_cell(row[src])

        if not out[_ITEM_KEY]:
            continue  # footer / subtotal / group-trailer row

        # --- wagon: cell -> id link (HR) -> load-plan last-5 vlookup (Plate) ---
        wagon = extract_wagon(row[wcol]) if wcol is not None else ""
        if not wagon and id_lookup:
            wagon = id_lookup.get(re.sub(r"\D", "", out["Batch Id."]), "")
        if not wagon and loadplan:
            # CONFIRM: the 5-digit key is the digits of the plate row's loading
            # code, read here from the wagon/loading column (wcol). If Plate keeps
            # that code in a differently-named column, point me at it.
            code = re.sub(r"\D", "", clean_cell(row[wcol])) if wcol is not None else ""
            wagon = loadplan.get(code[-5:], "")
        out["Wagon No."] = wagon
        failed += not wagon

        # --- refinements ---
        out["Pcs."] = "1"                                   # 2: always 1 per item
        weight = out["Act.ual Wt."] or out["Theoretical. Wt."]   # prefer PGI/actual
        out["Act.ual Wt."] = out["Theoretical. Wt."] = weight    # 3: both weights equal
        out["Material."] = classify_material(out)           # 5: material from dimensions
        if not out["Lot No."]:                              # 1: lot falls back to wagon
            out["Lot No."] = wagon

        rows.append(out)

    df_out = pd.DataFrame(rows, columns=output_schema).drop_duplicates(
        subset=[_ITEM_KEY], ignore_index=True)  # Batch Id is the unique primary key
    if vans:  # belt-and-suspenders: drop any item that resolved to a brake van
        keep = ~df_out["Wagon No."].isin(vans)
        dropped = int((~keep).sum())
        if dropped:
            logger(f"Dropped {dropped} item row(s) on brake-van wagons.")
        df_out = df_out[keep].reset_index(drop=True)
    df_out = df_out.sort_values("Wagon No.", kind="stable", ignore_index=True)  # 4: group wagons
    df_out["S.No."] = range(1, len(df_out) + 1)
    df_out.to_csv(output_path, index=False)

    matched = len(df_out) - failed
    logger(f"Wrote {len(df_out)} rows ({matched} with wagon, {failed} without).")
    return {
        "primary_sheet": detail_name,
        "total": len(df_out),
        "matched": matched,
        "fuzzy": 0,
        "failed": failed,
        "output_path": output_path,
    }


def _self_check() -> None:
    """Smallest checks that fail if the core logic breaks (ponytail: one test)."""
    assert clean_cell(607834255.0) == "607834255"  # no float tail on ids
    assert clean_cell(11.17) == "11.17"  # real decimals survive
    assert clean_cell(float("nan")) == ""
    assert extract_wagon("JSPL26452614091") == "26452614091"  # strip prefix
    assert extract_wagon(92452416232.0) == "92452416232"
    assert extract_wagon("JSPL10727") == ""  # 5-digit loading code, not a wagon
    cmap = build_colmap(pd.Index(["SOLD to Party", "Ship to Party", "THK.", "PGI WT", "NET WT"]))
    assert cmap["Sold To Party."] == "SOLD to Party"  # the old collision bug
    assert cmap["Ship To Party."] == "Ship to Party"
    assert cmap["Thickness."] == "THK."
    assert cmap["Act.ual Wt."] == "PGI WT" and cmap["Theoretical. Wt."] == "NET WT"

    # material from dimensions
    assert classify_material({"Diameter.": "12"}) == "TMT"
    assert classify_material({"Thickness.": "5", "Width.": "1250", "Length.": "6000"}) == "Plate"
    assert classify_material({"Thickness.": "2", "Width.": "1000"}) == "COIL"
    assert classify_material({"Material.": "SCRAP"}) == "SCRAP"  # fallback

    # load-plan last-5 index
    lp = build_loadplan_lookup({"Load Plan": pd.DataFrame({"x": ["JSPL26452610727"]})})
    assert lp == {"10727": "26452610727"}

    # WAGON column wins over VEHICLE NO (truck number)
    wc = wagon_column(pd.DataFrame(columns=["VEHICLE NO", "WAGON"]))
    assert wc == "WAGON"

    # brake vans (BV…) are detected and excluded
    assert _is_brake_van("BVCM") and _is_brake_van("bvzi") and _is_brake_van("BVZC")
    assert not _is_brake_van("BRN22.9") and not _is_brake_van("94452310527")
    bv_sheet = {"LOAD PLAN": pd.DataFrame(
        {"no": ["SER87072412313", "JSPL94452310459"], "type": ["BVCM", "BRN22.9"]})}
    assert brake_van_wagons(bv_sheet) == {"87072412313"}
    assert build_loadplan_lookup(bv_sheet) == {"10459": "94452310459"}  # van not indexed
    print("self-check OK")


if __name__ == "__main__":
    _self_check()
    sample = "Sample/PLATE-PGI DONE RAKE NO -21.xlsx"
    if Path(sample).exists():
        run_conversion(sample, SCHEMA, "output_cleaned.csv")
