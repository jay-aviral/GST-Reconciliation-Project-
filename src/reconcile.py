"""
GST Reconciliation: Purchase Register (Tally) vs GSTR-2B

Run from the repo root:
    python src/reconcile.py
"""
import re
import sys
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path

import pandas as pd

BOOKS_FILE = sys.argv[1] if len(sys.argv) > 1 else "data/purchase_register.xlsx"
GSTR2B_FILE = sys.argv[2] if len(sys.argv) > 2 else "data/gstr2b.xlsx"
OUT_FILE = sys.argv[3] if len(sys.argv) > 3 else "output/reconciliation_output.xlsx"

TAX_TOLERANCE = 2.0      # Rs. difference in total tax still treated as a match
VALUE_TOLERANCE = 1.0    # Rs. difference in taxable value for fuzzy rounds
SIMILARITY = 0.75        # invoice-number similarity needed in round 3
GSTIN_MAX_DIFF = 2       # max differing characters for a GSTIN typo


# ------------------------------------------------------------------ cleaning
def clean_inv(x) -> str:
    """Uppercase and keep only letters and digits: 'inv/0045 ' -> 'INV0045'."""
    return re.sub(r"[^A-Z0-9]", "", str(x).upper())


def clean_gstin(x) -> str:
    return str(x).strip().upper().replace(" ", "")


def parse_date(x):
    if isinstance(x, (datetime, pd.Timestamp)):
        return pd.Timestamp(x).date()
    s = str(x).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return pd.NaT


def prepare(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["GSTIN"] = df["GSTIN"].apply(clean_gstin)
    df["Inv_Clean"] = df["Invoice No"].apply(clean_inv)
    df["Invoice Date"] = df["Invoice Date"].apply(parse_date)
    for c in ["Taxable Value", "CGST", "SGST", "IGST"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    df["Total Tax"] = (df["CGST"] + df["SGST"] + df["IGST"]).round(2)
    df["Key"] = df["GSTIN"] + "|" + df["Inv_Clean"]
    return df.reset_index(drop=True)


def hamming(a: str, b: str) -> int:
    if len(a) != len(b):
        return 99
    return sum(x != y for x, y in zip(a, b))


# ------------------------------------------------------------------ load
books = prepare(pd.read_excel(BOOKS_FILE))
gst = prepare(pd.read_excel(GSTR2B_FILE))
print(f"Loaded books: {len(books)} rows | GSTR-2B: {len(gst)} rows")

# Flag duplicates in books (same GSTIN + invoice no + amount)
books["Is_Duplicate"] = books.duplicated(["Key", "Total Tax"], keep="first")

pairs = {}          # book idx -> (2B idx, status, remark)
used_2b = set()


def tax_status(b_idx, g_idx):
    diff = round(books.at[b_idx, "Total Tax"] - gst.at[g_idx, "Total Tax"], 2)
    if abs(diff) <= TAX_TOLERANCE:
        return "Matched"
    return "Value Mismatch"


# ------------------------------------------------------------------ round 1: exact key
g_by_key = {}
for gi, k in gst["Key"].items():
    g_by_key.setdefault(k, gi)

for bi, row in books.iterrows():
    if row["Is_Duplicate"]:
        continue
    gi = g_by_key.get(row["Key"])
    if gi is not None and gi not in used_2b:
        pairs[bi] = (gi, tax_status(bi, gi), "Exact key match")
        used_2b.add(gi)

# ------------------------------------------------------------------ round 2: GSTIN typo
rem_b = [bi for bi in books.index if bi not in pairs and not books.at[bi, "Is_Duplicate"]]
g_by_inv = {}
for gi in gst.index:
    if gi not in used_2b:
        g_by_inv.setdefault(gst.at[gi, "Inv_Clean"], []).append(gi)

for bi in rem_b:
    for gi in g_by_inv.get(books.at[bi, "Inv_Clean"], []):
        if gi in used_2b:
            continue
        if (abs(books.at[bi, "Taxable Value"] - gst.at[gi, "Taxable Value"]) <= VALUE_TOLERANCE
                and hamming(books.at[bi, "GSTIN"], gst.at[gi, "GSTIN"]) <= GSTIN_MAX_DIFF):
            pairs[bi] = (gi, "Probable Match (GSTIN typo)",
                         f"GSTIN in books {books.at[bi, 'GSTIN']} vs 2B {gst.at[gi, 'GSTIN']}")
            used_2b.add(gi)
            break

# ------------------------------------------------------------------ round 3: invoice no differs
rem_b = [bi for bi in books.index if bi not in pairs and not books.at[bi, "Is_Duplicate"]]
g_by_gstin = {}
for gi in gst.index:
    if gi not in used_2b:
        g_by_gstin.setdefault(gst.at[gi, "GSTIN"], []).append(gi)

for bi in rem_b:
    best, best_ratio = None, 0.0
    for gi in g_by_gstin.get(books.at[bi, "GSTIN"], []):
        if gi in used_2b:
            continue
        if abs(books.at[bi, "Taxable Value"] - gst.at[gi, "Taxable Value"]) > VALUE_TOLERANCE:
            continue
        r = SequenceMatcher(None, books.at[bi, "Inv_Clean"], gst.at[gi, "Inv_Clean"]).ratio()
        if r >= SIMILARITY and r > best_ratio:
            best, best_ratio = gi, r
    if best is not None:
        pairs[bi] = (best, "Probable Match (Invoice no differs)",
                     f"Books '{books.at[bi, 'Invoice No']}' vs 2B '{gst.at[best, 'Invoice No']}' "
                     f"(similarity {best_ratio:.2f})")
        used_2b.add(best)

# ------------------------------------------------------------------ build result
suppliers_in_2b = set(gst["GSTIN"])
rows = []
for bi, b in books.iterrows():
    base = {
        "Voucher No": b.get("Voucher No"), "GSTIN (Books)": b["GSTIN"],
        "Supplier": b["Supplier Name"], "Invoice No (Books)": b["Invoice No"],
        "Invoice Date": b["Invoice Date"], "Taxable (Books)": b["Taxable Value"],
        "Tax (Books)": b["Total Tax"],
    }
    if b["Is_Duplicate"]:
        status, remark, g = "Duplicate Entry in Books", "Same GSTIN, invoice no and amount booked twice", None
    elif bi in pairs:
        gi, status, remark = pairs[bi]
        g = gst.loc[gi]
    else:
        g = None
        status = "In Books Only (ITC at Risk)"
        remark = ("Supplier has no invoices in GSTR-2B (possible non-filer)"
                  if b["GSTIN"] not in suppliers_in_2b else "Invoice not reported by supplier in 2B")
    base.update({
        "GSTIN (2B)": g["GSTIN"] if g is not None else None,
        "Invoice No (2B)": g["Invoice No"] if g is not None else None,
        "Taxable (2B)": g["Taxable Value"] if g is not None else None,
        "Tax (2B)": g["Total Tax"] if g is not None else None,
        "Tax Difference": round(b["Total Tax"] - g["Total Tax"], 2) if g is not None else None,
        "Status": status, "Remark": remark,
    })
    rows.append(base)

for gi, g in gst.iterrows():
    if gi in used_2b:
        continue
    rows.append({
        "Voucher No": None, "GSTIN (Books)": None, "Supplier": g["Supplier Name"],
        "Invoice No (Books)": None, "Invoice Date": g["Invoice Date"],
        "Taxable (Books)": None, "Tax (Books)": None,
        "GSTIN (2B)": g["GSTIN"], "Invoice No (2B)": g["Invoice No"],
        "Taxable (2B)": g["Taxable Value"], "Tax (2B)": g["Total Tax"],
        "Tax Difference": None, "Status": "In 2B Only (Unrecorded Purchase)",
        "Remark": "Supplier reported it but it is not in books; check and record to claim ITC",
    })

result = pd.DataFrame(rows)

# ------------------------------------------------------------------ summaries
order = ["Matched", "Probable Match (GSTIN typo)", "Probable Match (Invoice no differs)",
         "Value Mismatch", "In Books Only (ITC at Risk)", "In 2B Only (Unrecorded Purchase)",
         "Duplicate Entry in Books"]
result["Tax (Amount)"] = result["Tax (Books)"].fillna(result["Tax (2B)"])
summary = (result.groupby("Status")
           .agg(Invoices=("Status", "size"), Tax_Amount=("Tax (Amount)", "sum"))
           .reindex(order).fillna(0))
summary["Invoices"] = summary["Invoices"].astype(int)
summary["Tax_Amount"] = summary["Tax_Amount"].round(2)
summary["% of Invoices"] = (summary["Invoices"] / summary["Invoices"].sum() * 100).round(1)

vm = result[result["Status"] == "Value Mismatch"]
excess_claimed = vm.loc[vm["Tax Difference"] > 0, "Tax Difference"].sum()
books_only_tax = result.loc[result["Status"] == "In Books Only (ITC at Risk)", "Tax (Books)"].sum()
itc_books = books.loc[~books["Is_Duplicate"], "Total Tax"].sum()
itc_2b = gst["Total Tax"].sum()
itc_at_risk = books_only_tax + excess_claimed
matched_n = summary.loc[["Matched", "Probable Match (GSTIN typo)",
                         "Probable Match (Invoice no differs)"], "Invoices"].sum()
match_pct = matched_n / (len(books) - books["Is_Duplicate"].sum()) * 100

kpis = pd.DataFrame({
    "Metric": ["Invoices in books (excl. duplicates)", "Invoices in GSTR-2B",
               "ITC as per books (Rs)", "ITC as per GSTR-2B (Rs)",
               "ITC at risk: missing in 2B (Rs)", "ITC at risk: books higher than 2B (Rs)",
               "Total ITC at risk (Rs)", "ITC at risk % of books ITC",
               "Unrecorded purchases ITC in 2B (Rs)", "Match rate (%)"],
    "Value": [int((~books["Is_Duplicate"]).sum()), len(gst), round(itc_books, 2), round(itc_2b, 2),
              round(books_only_tax, 2), round(excess_claimed, 2), round(itc_at_risk, 2),
              round(itc_at_risk / itc_books * 100, 2),
              round(result.loc[result["Status"] == "In 2B Only (Unrecorded Purchase)", "Tax (2B)"].sum(), 2),
              round(match_pct, 1)],
})

exc = result[result["Status"].isin(["In Books Only (ITC at Risk)", "Value Mismatch"])].copy()
exc["At Risk Amount"] = exc.apply(
    lambda r: r["Tax (Books)"] if r["Status"].startswith("In Books") else max(r["Tax Difference"], 0), axis=1)
supplier_risk = (exc.groupby(["GSTIN (Books)", "Supplier"])
                 .agg(Exception_Invoices=("Status", "size"), ITC_At_Risk=("At Risk Amount", "sum"))
                 .sort_values("ITC_At_Risk", ascending=False).reset_index())
supplier_risk["ITC_At_Risk"] = supplier_risk["ITC_At_Risk"].round(2)
supplier_risk["Cumulative %"] = (supplier_risk["ITC_At_Risk"].cumsum()
                                 / supplier_risk["ITC_At_Risk"].sum() * 100).round(1)

# ------------------------------------------------------------------ write Excel
Path(OUT_FILE).parent.mkdir(parents=True, exist_ok=True)
with pd.ExcelWriter(OUT_FILE, engine="openpyxl") as xw:
    kpis.to_excel(xw, sheet_name="KPIs", index=False)
    summary.reset_index().to_excel(xw, sheet_name="Status Summary", index=False)
    supplier_risk.to_excel(xw, sheet_name="Supplier Risk", index=False)
    result.drop(columns="Tax (Amount)").to_excel(xw, sheet_name="Reconciliation", index=False)
    exc.drop(columns="Tax (Amount)").to_excel(xw, sheet_name="Exceptions", index=False)
    for ws in xw.book.worksheets:
        ws.freeze_panes = "A2"
        for col in ws.columns:
            width = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200])
            ws.column_dimensions[col[0].column_letter].width = min(width + 2, 45)

print("\n=== STATUS SUMMARY ===")
print(summary.to_string())
print("\n=== KPIs ===")
print(kpis.to_string(index=False))
print("\n=== TOP 5 SUPPLIERS BY ITC AT RISK ===")
print(supplier_risk.head(5).to_string(index=False))
print(f"\nSaved: {OUT_FILE}")
