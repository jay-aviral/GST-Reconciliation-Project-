"""
GST Reconciliation Dashboard (Streamlit)

Run from the repo root:
    pip install streamlit plotly pandas openpyxl
    streamlit run src/dashboard.py
"""
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

DATA_FILE = Path("output/reconciliation_output.xlsx")

MATCHED = ["Matched", "Probable Match (GSTIN typo)", "Probable Match (Invoice no differs)"]
BOOKS_ONLY = "In Books Only (ITC at Risk)"
ONLY_2B = "In 2B Only (Unrecorded Purchase)"
DUPLICATE = "Duplicate Entry in Books"
MISMATCH = "Value Mismatch"

COLORS = {
    "Matched": "#2E8B57",
    "Probable Match (GSTIN typo)": "#7FBF7F",
    "Probable Match (Invoice no differs)": "#B7DDB0",
    MISMATCH: "#F2A541",
    BOOKS_ONLY: "#D64545",
    ONLY_2B: "#4C78C8",
    DUPLICATE: "#8C8C8C",
}

st.set_page_config(page_title="GST Reconciliation Dashboard", page_icon="🧾", layout="wide")


def inr(x: float) -> str:
    """Format a rupee amount in lakh / crore."""
    x = float(x)
    if abs(x) >= 1e7:
        return f"₹{x / 1e7:.2f} Cr"
    if abs(x) >= 1e5:
        return f"₹{x / 1e5:.1f} L"
    return f"₹{x:,.0f}"


@st.cache_data
def load(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Reconciliation")
    df["Invoice Date"] = pd.to_datetime(df["Invoice Date"], errors="coerce")
    df["Month"] = df["Invoice Date"].dt.to_period("M").dt.to_timestamp()
    df["Supplier"] = df["Supplier"].fillna("Unknown")

    def at_risk(r):
        if r["Status"] == BOOKS_ONLY:
            return r["Tax (Books)"]
        if r["Status"] == MISMATCH:
            return max(r["Tax Difference"], 0)
        return 0.0

    df["ITC At Risk"] = df.apply(at_risk, axis=1)
    return df


if not DATA_FILE.exists():
    st.error(f"`{DATA_FILE}` not found. Run `python src/reconcile.py` first.")
    st.stop()

df = load(DATA_FILE)

# ------------------------------------------------------------------ sidebar filters
st.sidebar.header("Filters")
months = sorted(df["Month"].dropna().unique())
month_labels = {m: pd.Timestamp(m).strftime("%b %Y") for m in months}
sel_months = st.sidebar.multiselect("Month", months, default=months, format_func=lambda m: month_labels[m])
sel_status = st.sidebar.multiselect("Status", sorted(df["Status"].unique()), default=sorted(df["Status"].unique()))
sel_suppliers = st.sidebar.multiselect("Supplier (blank = all)", sorted(df["Supplier"].unique()))

f = df[df["Month"].isin(sel_months) & df["Status"].isin(sel_status)]
if sel_suppliers:
    f = f[f["Supplier"].isin(sel_suppliers)]

# ------------------------------------------------------------------ KPIs
books_rows = f[~f["Status"].isin([ONLY_2B, DUPLICATE])]
itc_books = f.loc[f["Status"] != DUPLICATE, "Tax (Books)"].sum()
itc_2b = f["Tax (2B)"].sum()
itc_risk = f["ITC At Risk"].sum()
match_rate = (books_rows["Status"].isin(MATCHED).sum() / len(books_rows) * 100) if len(books_rows) else 0
unrecorded = f.loc[f["Status"] == ONLY_2B, "Tax (2B)"].sum()
dup_itc = f.loc[f["Status"] == DUPLICATE, "Tax (Books)"].sum()

st.title("🧾 GST Reconciliation Dashboard")
st.caption("Purchase Register (Tally) vs GSTR-2B | synthetic data")

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("ITC as per Books", inr(itc_books))
c2.metric("ITC as per GSTR-2B", inr(itc_2b))
c3.metric("ITC at Risk", inr(itc_risk), f"{itc_risk / itc_books * 100:.1f}% of books ITC" if itc_books else None,
          delta_color="inverse")
c4.metric("Match Rate", f"{match_rate:.1f}%")
c5.metric("Unrecorded in Books", inr(unrecorded))

if dup_itc:
    st.warning(f"{inr(dup_itc)} of ITC is from duplicate entries in books and should be removed.")

# ------------------------------------------------------------------ charts row 1
left, right = st.columns(2)

with left:
    st.subheader("Invoices by Status")
    s = f.groupby("Status").size().reset_index(name="Invoices")
    fig = px.pie(s, names="Status", values="Invoices", hole=0.55,
                 color="Status", color_discrete_map=COLORS)
    fig.update_traces(textinfo="percent", sort=False)
    fig.update_layout(margin=dict(t=10, b=10), legend=dict(orientation="h", y=-0.2))
    st.plotly_chart(fig, use_container_width=True)

with right:
    st.subheader("Top 10 Suppliers by ITC at Risk")
    top = (f[f["ITC At Risk"] > 0].groupby("Supplier")["ITC At Risk"].sum()
           .sort_values(ascending=False).head(10).reset_index())
    if top.empty:
        st.info("No ITC at risk for the selected filters.")
    else:
        fig = px.bar(top.sort_values("ITC At Risk"), x="ITC At Risk", y="Supplier", orientation="h",
                     color_discrete_sequence=["#D64545"])
        fig.update_layout(margin=dict(t=10, b=10), xaxis_title="ITC at risk (₹)", yaxis_title=None)
        st.plotly_chart(fig, use_container_width=True)

# ------------------------------------------------------------------ charts row 2
st.subheader("Monthly ITC: Books vs GSTR-2B vs At Risk")
m = f.groupby("Month").agg(Books=("Tax (Books)", "sum"), GSTR2B=("Tax (2B)", "sum"),
                           AtRisk=("ITC At Risk", "sum")).reset_index()
m_long = m.melt("Month", var_name="Measure", value_name="Amount")
fig = px.bar(m_long, x="Month", y="Amount", color="Measure", barmode="group",
             color_discrete_map={"Books": "#4C78C8", "GSTR2B": "#2E8B57", "AtRisk": "#D64545"})
fig.update_layout(margin=dict(t=10, b=10), yaxis_title="Tax amount (₹)", xaxis_title=None)
st.plotly_chart(fig, use_container_width=True)

# ------------------------------------------------------------------ tables
tab1, tab2 = st.tabs(["Exceptions", "Supplier summary"])

with tab1:
    exc = f[f["Status"].isin([BOOKS_ONLY, MISMATCH, ONLY_2B, DUPLICATE])]
    cols = ["Voucher No", "Supplier", "Invoice No (Books)", "Invoice No (2B)", "Invoice Date",
            "Tax (Books)", "Tax (2B)", "Tax Difference", "ITC At Risk", "Status", "Remark"]
    st.dataframe(exc[cols].sort_values("ITC At Risk", ascending=False), use_container_width=True, height=380)
    st.download_button("Download exceptions (CSV)", exc[cols].to_csv(index=False).encode("utf-8"),
                       file_name="gst_exceptions.csv", mime="text/csv")

with tab2:
    sup = (f.groupby("Supplier")
           .agg(Invoices=("Status", "size"),
                Exceptions=("Status", lambda x: (~x.isin(MATCHED)).sum()),
                ITC_At_Risk=("ITC At Risk", "sum"))
           .sort_values("ITC_At_Risk", ascending=False).reset_index())
    sup["Exception %"] = (sup["Exceptions"] / sup["Invoices"] * 100).round(1)
    st.dataframe(sup, use_container_width=True, height=380)