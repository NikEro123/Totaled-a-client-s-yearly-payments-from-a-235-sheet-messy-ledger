
import os

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
OUTPUT = os.path.join(HERE, "borjomi_xelfasi_2025.xlsx")

YEAR = 2025
CLIENT = "borjomi"                     
CLIENT_GE = "ბორჯომი"                   
PAYMENT_WORDS = ("xelfasi", "ხელფასი")
LABELS = {"tariri", "tanxa", "semosavlis wyaro", "daniSnuleba".lower(),
          "jami", "naSti".lower(), "naSTi".lower(), "wina naSti".lower()}

notes = []


def find_file(start, end):
    for name in os.listdir(HERE):
        if name.lower().startswith(start.lower()) and name.lower().endswith(end.lower()):
            return os.path.join(HERE, name)
    raise FileNotFoundError(f"No file starting with '{start}' and ending with '{end}' in {HERE}")


def text(value):
    """Cell as clean text ('' if empty)."""
    if pd.isna(value):
        return ""
    return " ".join(str(value).split())


def number(value):
    """Cell as a number, or None. Handles '120', '120,5', 120."""
    if pd.isna(value):
        return None
    try:
        return float(str(value).replace(" ", "").replace(",", "."))
    except ValueError:
        return None


def sheet_period(name):
    """'30,02,25 (2)' -> (2025, 2).  '01-30,01,25' -> (2025, 1).  '13.12' -> (None, 12)."""
    base = name.split(" (")[0].replace(".", ",")
    parts = [p for p in base.split(",") if p.strip()]
    month = year = None
    if len(parts) >= 2 and parts[1].strip().isdigit():
        month = int(parts[1])
    if len(parts) >= 3:
        y = parts[-1].strip()[-2:]
        if y.isdigit():
            year = 2000 + int(y)
    return year, month


def entry_date(value):
    """Ledger date cell -> (day, month, year or None). Accepts real dates and '03,01,25'."""
    if hasattr(value, "year") and hasattr(value, "month"):
        return value.day, value.month, value.year
    parts = [p for p in text(value).replace(".", ",").split(",") if p.strip()]
    if len(parts) >= 2 and parts[0].isdigit() and parts[1].isdigit():
        year = 2000 + int(parts[2][-2:]) if len(parts) >= 3 and parts[2][-2:].isdigit() else None
        return int(parts[0]), int(parts[1]), year
    return None, None, None


def read_entries(sheet_name, df):
    """Walk one ledger sheet and return every xelfasi payment with its client."""
    sheet_year, sheet_month = sheet_period(sheet_name)
    entries = []
    client, date_cell, waiting_for_amount = None, None, False

    for _, row in df.iterrows():
        a, b, c = text(row.iloc[0]), text(row.iloc[1]).lower(), number(row.iloc[2])
        e = row.iloc[4] if len(row) > 4 else None

        if a:                                   
            if a.lower() in LABELS:
                client, waiting_for_amount = None, False  
                continue
            client, date_cell, waiting_for_amount = a, e, False

        if client is None:
            continue

        is_payment = any(word in b for word in PAYMENT_WORDS)
        amount = None
        if is_payment and c is not None:
            amount = c
        elif is_payment:
            waiting_for_amount = True           
        elif waiting_for_amount and not b and c is not None:
            amount, waiting_for_amount = c, False

        if amount is not None:
            day, month, year = entry_date(date_cell)
            final_year, final_month = year or sheet_year, month or sheet_month
            if year and sheet_year and year != sheet_year and month == sheet_month:
                notes.append(f"{sheet_name}: '{text(date_cell)}' looks like a year typo -> used {sheet_year}")
                final_year = sheet_year
            entries.append({
                "client": client, "amount": amount, "date_as_typed": text(date_cell),
                "year": final_year, "month": final_month, "sheet": sheet_name,
            })
    return entries


def main():
    path = find_file("bugelteria", ".xlsx")
    print(f"Reading {os.path.basename(path)} ...")
    sheets = pd.read_excel(path, sheet_name=None, header=None)

    rows = []
    for name, df in sheets.items():
        rows.extend(read_entries(name, df))
    all_payments = pd.DataFrame(rows)

    is_borjomi = all_payments["client"].str.lower().str.contains(CLIENT) | \
        all_payments["client"].str.contains(CLIENT_GE)
    borjomi = all_payments[is_borjomi & (all_payments["year"] == YEAR)].copy()
    borjomi = borjomi.sort_values(["month", "sheet"]).reset_index(drop=True)

    by_month = (borjomi.groupby("month")
                .agg(payments=("amount", "size"), total=("amount", "sum"))
                .reindex(range(1, 13), fill_value=0).reset_index())
    by_month["month"] = pd.to_datetime(by_month["month"], format="%m").dt.strftime("%B")

    total = borjomi["amount"].sum()
    summary = pd.DataFrame({
        "metric": ["Client", "Year", "Number of xelfasi payments", "Total xelfasi", "Average payment",
                   "Smallest payment", "Largest payment"],
        "value": ["BORJOMI", YEAR, len(borjomi), total, round(borjomi["amount"].mean(), 2),
                  borjomi["amount"].min(), borjomi["amount"].max()],
    })

    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        by_month.to_excel(writer, sheet_name="By Month", index=False)
        borjomi.to_excel(writer, sheet_name="All Entries", index=False)
        pd.DataFrame({"note": notes or ["No date typos found"]}).to_excel(writer, sheet_name="Notes", index=False)
        for ws in writer.sheets.values():
            ws.freeze_panes = "A2"
            for col_cells in ws.columns:
                width = max(len(str(c.value or "")) for c in col_cells[:200]) + 2
                ws.column_dimensions[col_cells[0].column_letter].width = min(max(width, 10), 60)

    print(f"\nBORJOMI xelfasi in {YEAR}: {len(borjomi)} payments, total {total:,.2f}")
    print(by_month.to_string(index=False))
    if notes:
        print(f"\n{len(notes)} date typos handled (see the Notes sheet)")
    print(f"\nSaved: {OUTPUT}")


if __name__ == "__main__":
    main()
