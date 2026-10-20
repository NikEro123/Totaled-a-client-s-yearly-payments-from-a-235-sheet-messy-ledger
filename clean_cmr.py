

import os

import numpy as np
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
OUTPUT = os.path.join(HERE, "cmr_clean.xlsx")

FIELDS = {
    "sender":          (1, 0),
    "receiver":        (7, 0),
    "carrier":         (7, 3),
    "unload_place":    (13, 0),
    "load_place_date": (17, 0),
    "container_raw":   (25, 0),
    "packages":        (25, 1),
    "goods":           (25, 2),
    "weight_raw":      (25, 5),
    "doc_place":       (50, 0),
    "doc_date_raw":    (50, 2),
    "driver_raw":      (53, 2),
    "truck_raw":       (58, 0),
    "truck_make":      (58, 2),
}


GEORGIAN_FONT = {
    "a": "ა", "b": "ბ", "g": "გ", "d": "დ", "e": "ე", "v": "ვ", "z": "ზ", "T": "თ",
    "i": "ი", "k": "კ", "l": "ლ", "m": "მ", "n": "ნ", "o": "ო", "p": "პ", "J": "ჟ",
    "r": "რ", "s": "ს", "t": "ტ", "u": "უ", "f": "ფ", "q": "ქ", "R": "ღ", "y": "ყ",
    "S": "შ", "C": "ჩ", "c": "ც", "Z": "ძ", "w": "წ", "W": "ჭ", "x": "ხ", "j": "ჯ",
    "h": "ჰ",
}

log = []


def note(fix, rows):
    """Add one line to the Data Quality report (lines with 0 rows are skipped)."""
    if int(rows) > 0:
        log.append({"fix": fix, "rows": int(rows)})



def find_file(start, end):
    for name in os.listdir(HERE):
        if name.lower().startswith(start.lower()) and name.lower().endswith(end.lower()):
            return os.path.join(HERE, name)
    raise FileNotFoundError(f"No file starting with '{start}' and ending with '{end}' in {HERE}")


def read_cell(sheet, row, col):
    if row >= sheet.shape[0] or col >= sheet.shape[1]:
        return None
    value = sheet.iloc[row, col]
    if pd.isna(value):
        return None
    if isinstance(value, str):
        value = " ".join(value.split())
        return value or None
    return value


def extract(path):
    sheets = pd.read_excel(path, sheet_name=None, header=None)
    rows = []
    for sheet_name, sheet in sheets.items():
        if sheet.empty:
            continue
        row = {"sheet": sheet_name}
        for field, (r, c) in FIELDS.items():
            row[field] = read_cell(sheet, r, c)
        rows.append(row)
    note("Empty sheets skipped", len(sheets) - len(rows))
    return pd.DataFrame(rows)



def to_georgian(text):
    """Convert old-font words to real Georgian, word by word.

    A word is converted only if it has a lowercase Latin letter and no digits:
      - "foTi" -> "ფოთი"   (old font)
      - "MAXTHERM", "LTD", "CMAU8493382" -> unchanged (real Latin / codes)
      - "ფოთი" -> unchanged (already real Georgian)
    """
    if not isinstance(text, str):
        return text
    words = []
    for word in text.split(" "):
        has_lower = any("a" <= ch <= "z" for ch in word)
        has_digit = any(ch.isdigit() for ch in word)
        if has_lower and not has_digit:
            word = "".join(GEORGIAN_FONT.get(ch, ch) for ch in word)
        words.append(word)
    return " ".join(words)


def parse_dates(text_series, default_year=None):
    """Parse dates like '20,03,26', '09.02.15', '03.06,25', '04.10' (no year).

    default_year: a Series of years to use when the date has no year.
    """
    parts = text_series.astype("string").str.extract(r"(\d{1,2})\s*[.,]+\s*(\d{1,2})(?:\s*[.,]+\s*(\d{2,4}))?")
    day = pd.to_numeric(parts[0], errors="coerce")
    month = pd.to_numeric(parts[1], errors="coerce")
    year = pd.to_numeric(parts[2].str[-2:], errors="coerce").astype("float64") + 2000   # '021' typo -> 21 -> 2021
    if default_year is not None:
        year = year.fillna(pd.Series(default_year, index=text_series.index).astype("float64"))
    built = pd.DataFrame({"year": year, "month": month, "day": day}).astype("float64")
    ok = built.notna().all(axis=1)
    result = pd.Series(pd.NaT, index=text_series.index, dtype="datetime64[ns]")
    result[ok] = pd.to_datetime(built[ok], errors="coerce")   # impossible dates (31.02) -> NaT
    return result



def clean(df):
    total = len(df)

    # 1. Old Georgian font -> real Georgian
    text_cols = ["sender", "receiver", "carrier", "unload_place", "goods", "doc_place", "truck_make"]
    converted = 0
    for col in text_cols:
        new = df[col].map(to_georgian)
        converted += (new.astype("string") != df[col].astype("string")).fillna(False).sum()
        df[col] = new
    note("Text cells converted from old Georgian font", converted)


    raw = df["doc_date_raw"]
    is_dt = raw.map(lambda v: hasattr(v, "year") and hasattr(v, "month"))   
    is_num = raw.map(lambda v: isinstance(v, (int, float, np.number)) and not isinstance(v, bool))
    doc_date = pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
    doc_date[is_dt] = pd.to_datetime(raw[is_dt])
 
  
    typed_as_number = is_num & (raw.where(is_num, 0).astype(float) < 100)
    is_serial = is_num & ~typed_as_number
    doc_date[is_serial] = pd.to_datetime(raw[is_serial].astype(float), unit="D", origin="1899-12-30")
    as_text = raw.where(~typed_as_number, raw.map(lambda v: f"{v:.2f}" if isinstance(v, (int, float)) else v))
    is_text = ~(is_dt | is_serial) & raw.notna()
    doc_date[is_text] = parse_dates(as_text[is_text])
    note("Document dates read as Excel dates", is_dt.sum() + is_serial.sum())
    note("Document dates converted from text (e.g. '20,03,26')", doc_date[is_text].notna().sum())
    note("Dates typed as numbers (e.g. 6.08 = 6 August) fixed", typed_as_number.sum())


    lpd = df["load_place_date"].astype("string")
    load_date = parse_dates(lpd, default_year=doc_date.dt.year)
    df["load_place"] = lpd.str.replace(r"\d{1,2}\s*[.,]+\s*\d{1,2}(\s*[.,]+\s*\d{2,4})?", "", regex=True)
    df["load_place"] = df["load_place"].str.strip(" ,.").map(to_georgian)
    note("Loading dates split out of the place text", load_date.notna().sum())


    impossible = (doc_date.dt.year < 2010) | (doc_date > pd.Timestamp.today())
    note("Impossible document dates (e.g. year 2054) replaced by loading date", impossible.sum())
    doc_date[impossible] = pd.NaT


    df["date"] = doc_date.fillna(load_date)
    note("Missing document date filled with loading date", (doc_date.isna() & load_date.notna()).sum())
    note("CMRs with no usable date at all", df["date"].isna().sum())


  
    cont = df["container_raw"].astype("string")
    df["container_size_ft"] = pd.to_numeric(cont.str.extract(r"\d\s*[xXхХ]\s*(20|40|45)")[0], errors="coerce")
    num = cont.str.upper().str.extract(r"([A-Z]{4})\s*(\d{6,7})")
    df["container_no"] = num[0] + num[1]
    df["seal_no"] = cont.str.extract(r"[SsSs]\s*[:.]\s*(\d{5,})")[0]
    df["empty_container"] = df["goods"].isin(["ცარიელი", "carieli"])
    note("Container numbers extracted (e.g. TIIU5701646)", df["container_no"].notna().sum())
    note("CMRs with no container number found", df["container_no"].isna().sum())

    
    df["packages"] = pd.to_numeric(df["packages"], errors="coerce")
    w = df["weight_raw"].astype("string").str.replace(" ", "", regex=False).str.replace(",", ".", regex=False)
    df["weight_kg"] = pd.to_numeric(w, errors="coerce")
    fixed_w = df["weight_raw"].map(lambda v: isinstance(v, str)) & df["weight_kg"].notna()
    note("Weights fixed from text (e.g. '26996, 95')", fixed_w.sum())

  
    drv = df["driver_raw"].astype("string")
    had_id = drv.str.contains(r"\d{9,}", regex=True).fillna(False)
    df["driver"] = drv.str.split(r"\d{9,}", n=1, regex=True).str[0].str.strip().map(to_georgian)
    note("Driver personal ID numbers removed", had_id.sum())


    plates = df["truck_raw"].astype("string").str.upper().str.split("/", n=1, expand=True)
    df["truck_plate"] = plates[0].str.replace(" ", "", regex=False)
    df["trailer_plate"] = plates[1].str.replace(" ", "", regex=False) if 1 in plates else pd.NA


    key_cols = ["date", "container_no", "truck_plate", "weight_kg"]
    df["possible_duplicate"] = df.duplicated(subset=key_cols, keep=False) & df["container_no"].notna()
    note("CMRs flagged as possible duplicates", df["possible_duplicate"].sum())
    note("CMRs in the clean table", total)

    columns = ["date", "sender", "receiver", "carrier", "load_place", "unload_place",
               "container_size_ft", "container_no", "seal_no", "goods", "empty_container",
               "packages", "weight_kg", "driver", "truck_plate", "trailer_plate", "truck_make",
               "doc_place", "possible_duplicate", "sheet"]
    return df[columns].sort_values("date", na_position="last").reset_index(drop=True)


def save(clean_df, out):
    with pd.ExcelWriter(out, engine="openpyxl", datetime_format="YYYY-MM-DD") as writer:
        clean_df.to_excel(writer, sheet_name="CMRs", index=False)
        pd.DataFrame(log).to_excel(writer, sheet_name="Data Quality", index=False)
        for ws in writer.sheets.values():
            ws.freeze_panes = "A2"
            for col_cells in ws.columns:
                width = max(len(str(c.value or "")) for c in col_cells[:200]) + 2
                ws.column_dimensions[col_cells[0].column_letter].width = min(max(width, 10), 45)


def main():
    path = find_file("IEL", "2024.xls")
    print(f"Reading {os.path.basename(path)} ...")
    raw = extract(path)
    clean_df = clean(raw)
    save(clean_df, OUTPUT)

    print(pd.DataFrame(log).to_string(index=False))
    print(f"\nDates from {clean_df['date'].min():%Y-%m-%d} to {clean_df['date'].max():%Y-%m-%d}")
    print(f"Saved: {OUTPUT}")


if __name__ == "__main__":
    main()
