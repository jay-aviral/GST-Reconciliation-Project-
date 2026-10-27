"""
Synthetic data generator for GST Reconciliation project
(Purchase Register vs GSTR-2B). All data is FAKE.
Run from the data/ folder: python ../src/generate_gst_data.py
Outputs: purchase_register.xlsx, gstr2b.xlsx, answer_key.xlsx
"""
import random
import string
from datetime import date, timedelta

import pandas as pd

random.seed(42)

BUYER_STATE = "29"  # Karnataka (buyer in Bangalore)
N_INVOICES = 1500
START, END = date(2026, 4, 1), date(2026, 9, 30)

STATES = {"29": "Karnataka", "33": "Tamil Nadu", "27": "Maharashtra",
          "07": "Delhi", "36": "Telangana", "24": "Gujarat", "19": "West Bengal"}

NAMES = [
    "Sri Lakshmi Traders", "Metro Steel Supplies", "Apex Electricals", "Bharat Cement Depot",
    "Kaveri Hardware", "Sunrise Packaging", "Orion Office Systems", "Greenfield Paints",
    "Vijay Plumbing Mart", "National Tools & Dies", "Shree Ganesh Timber", "Deccan Logistics",
    "Prime Safety Gear", "Alpha IT Solutions", "Ramesh Stationers", "Blue Star Fabricators",
    "Maruti Aggregates", "City Glass House", "Eastern Cables Ltd", "Universal Lubricants",
    "Pioneer Tiles", "Southern Bricks Co", "Krishna Machinery", "Royal Fasteners",
    "Zenith Waterproofing", "Omkar Scaffolding", "Silverline Transport", "Vishwa Roofing",
    "Techno Power Systems", "Anand Sanitaryware", "Galaxy Housekeeping", "Mahalaxmi Enterprises",
    "Sagar Chemicals", "Ultra Fire Safety", "Nandi Granites", "Classic Furnitures",
    "Delta Pumps", "Everest Rebar Traders", "Falcon Security Services", "Heritage Marble",
]

GST_RATES = [5, 12, 18, 28]
CHARS = string.digits + string.ascii_uppercase


def gstin_checksum(g14: str) -> str:
    factor, total = 1, 0
    for ch in g14:
        v = CHARS.index(ch) * factor
        total += v // 36 + v % 36
        factor = 2 if factor == 1 else 1
    return CHARS[(36 - total % 36) % 36]


def make_gstin(state: str) -> str:
    pan = ("".join(random.choices(string.ascii_uppercase, k=5))
           + "".join(random.choices(string.digits, k=4))
           + random.choice(string.ascii_uppercase))
    base = state + pan + "1Z"
    return base + gstin_checksum(base)


# ---- suppliers ----
suppliers = []
for n in NAMES:
    st = random.choices(list(STATES), weights=[55, 12, 10, 6, 7, 5, 5])[0]
    suppliers.append({"name": n, "state": st, "gstin": make_gstin(st)})

# 3 suppliers who never filed GSTR-1 -> none of their invoices appear in 2B
non_filers = {s["gstin"] for s in random.sample(suppliers, 3)}

# ---- true invoices ----
rows = []
for i in range(N_INVOICES):
    s = random.choice(suppliers)
    d = START + timedelta(days=random.randint(0, (END - START).days))
    taxable = round(random.choice([
        random.uniform(2_000, 25_000), random.uniform(25_000, 150_000),
        random.uniform(150_000, 600_000)]), 2)
    if random.random() < 0.15:  # some round-figure invoices
        taxable = float(round(taxable, -3)) or 1000.0
    rate = random.choice(GST_RATES)
    tax = round(taxable * rate / 100, 2)
    intra = s["state"] == BUYER_STATE
    prefix = random.choice(["INV/", "INV-", "BILL", "SI/26-27/", "TI"])
    no = f"{prefix}{random.randint(1, 9999):04d}"
    rows.append({
        "GSTIN": s["gstin"], "Supplier Name": s["name"], "Invoice No": no,
        "Invoice Date": d, "Taxable Value": taxable,
        "CGST": round(tax / 2, 2) if intra else 0.0,
        "SGST": round(tax / 2, 2) if intra else 0.0,
        "IGST": 0.0 if intra else tax,
        "Rate": rate,
    })
true_df = pd.DataFrame(rows)
true_df = true_df.drop_duplicates(["GSTIN", "Invoice No"]).reset_index(drop=True)
true_df["Ref"] = ["R%05d" % (i + 1) for i in range(len(true_df))]

books = true_df.copy()
gst2b = true_df.copy()
books["Issue"] = "Matched"

# ---- plant errors (by Ref so we can verify later) ----
refs = list(true_df["Ref"])
random.shuffle(refs)
pool = iter(refs)
take = lambda k: [next(pool) for _ in range(k)]
n = len(true_df)

# 1) supplier never filed -> missing from 2B
nf_refs = set(true_df.loc[true_df["GSTIN"].isin(non_filers), "Ref"])
gst2b = gst2b[~gst2b["Ref"].isin(nf_refs)]
books.loc[books["Ref"].isin(nf_refs), "Issue"] = "In Books Only - supplier not filed"

# 2) random invoices missing in 2B (late filing)
miss = [r for r in take(int(n * 0.04)) if r not in nf_refs]
gst2b = gst2b[~gst2b["Ref"].isin(miss)]
books.loc[books["Ref"].isin(miss), "Issue"] = "In Books Only - not reported"

# 3) in 2B but not recorded in books
only2b = [r for r in take(int(n * 0.03)) if r not in nf_refs and r not in miss]
books = books[~books["Ref"].isin(only2b)]
gst2b.loc[gst2b["Ref"].isin(only2b), "Issue"] = "In 2B Only - unrecorded purchase"

# 4) value mismatch (supplier reported a different amount)
mism = [r for r in take(int(n * 0.04)) if r not in nf_refs and r not in miss and r not in only2b]
for r in mism:
    m = gst2b["Ref"] == r
    if not m.any():
        continue
    f = random.choice([0.9, 0.95, 1.05, 1.1])
    for c in ["Taxable Value", "CGST", "SGST", "IGST"]:
        gst2b.loc[m, c] = (gst2b.loc[m, c] * f).round(2)
    books.loc[books["Ref"] == r, "Issue"] = "Value Mismatch"

# 5) tiny rounding differences (should still MATCH within tolerance)
small = [r for r in take(int(n * 0.03)) if r in set(gst2b["Ref"]) and r in set(books["Ref"])
         and books.loc[books["Ref"] == r, "Issue"].iloc[0] == "Matched"]
for r in small:
    m = gst2b["Ref"] == r
    col = "IGST" if gst2b.loc[m, "IGST"].iloc[0] > 0 else "CGST"
    gst2b.loc[m, col] = (gst2b.loc[m, col] + random.choice([-1, 1, 0.5, -0.5])).round(2)
    books.loc[books["Ref"] == r, "Issue"] = "Matched (rounding diff)"

# 6) invoice-number format differences in books (fix by cleaning)
def mangle(no: str) -> str:
    return random.choice([
        no.lower(), no.replace("/", "").replace("-", ""),
        no.replace("INV", "INV ").replace("BILL", "BILL "),
        no.replace("0", "", 1) if "0" in no else no + " ",
        " " + no])
fmt = [r for r in take(int(n * 0.07)) if r in set(books["Ref"])]
for r in fmt:
    m = books["Ref"] == r
    books.loc[m, "Invoice No"] = mangle(books.loc[m, "Invoice No"].iloc[0])
    if books.loc[m, "Issue"].iloc[0] == "Matched":
        books.loc[m, "Issue"] = "Matched (invoice no format differs)"

# 7) typo in GSTIN in books
typo = [r for r in take(int(n * 0.01)) if r in set(books["Ref"])
        and books.loc[books["Ref"] == r, "Issue"].iloc[0] == "Matched"]
for r in typo:
    m = books["Ref"] == r
    g = books.loc[m, "GSTIN"].iloc[0]
    pos = random.randint(2, 11)
    books.loc[m, "GSTIN"] = g[:pos] + random.choice([c for c in string.ascii_uppercase if c != g[pos]]) + g[pos + 1:]
    books.loc[m, "Issue"] = "Wrong GSTIN in books"

# 8) duplicate entries in books
dups = books.sample(int(len(books) * 0.01), random_state=1).copy()
dups["Issue"] = "Duplicate entry in books"
books = pd.concat([books, dups], ignore_index=True)

# 9) messy formatting in books: lowercase/space GSTIN, mixed date formats
def messy_gstin(g):
    x = random.random()
    return g.lower() if x < 0.02 else (g + " " if x < 0.04 else g)
books["GSTIN"] = books["GSTIN"].apply(messy_gstin)
books["Invoice Date"] = books["Invoice Date"].apply(
    lambda d: d.strftime(random.choice(["%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"])) if random.random() < 0.3 else d)

# ---- finalize ----
books = books.sort_values("Ref").reset_index(drop=True)
gst2b = gst2b.sort_values("Ref").reset_index(drop=True)

books_out = books[["GSTIN", "Supplier Name", "Invoice No", "Invoice Date",
                   "Taxable Value", "CGST", "SGST", "IGST"]].copy()
books_out.insert(0, "Voucher No", ["PV%05d" % (i + 1) for i in range(len(books_out))])

gst2b_out = gst2b[["GSTIN", "Supplier Name", "Invoice No", "Invoice Date",
                   "Taxable Value", "CGST", "SGST", "IGST"]].copy()
gst2b_out["Invoice Date"] = pd.to_datetime(gst2b_out["Invoice Date"]).dt.date
gst2b_out["ITC Available"] = "Yes"
gst2b_out["Return Period"] = pd.to_datetime(gst2b_out["Invoice Date"]).dt.strftime("%b-%Y")

key = pd.concat([books[["Ref", "GSTIN", "Invoice No", "Issue"]].assign(Source="Books"),
                 gst2b[gst2b["Issue"] == "In 2B Only - unrecorded purchase"]
                 [["Ref", "GSTIN", "Invoice No", "Issue"]].assign(Source="2B")])

books_out.to_excel("purchase_register.xlsx", index=False)
gst2b_out.to_excel("gstr2b.xlsx", index=False)
key.to_excel("answer_key.xlsx", index=False)

print("Books rows:", len(books_out), "| 2B rows:", len(gst2b_out))
print(key["Issue"].value_counts())