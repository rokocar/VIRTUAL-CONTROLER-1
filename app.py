import streamlit as st
import pandas as pd
import numpy as np
import datetime as dt

st.set_page_config(page_title="Virtual Controller — Auto (Two Files)", layout="wide")
st.title("Virtual Controller — Auto (Two Files)")

st.markdown(
    "Naloži **dve Excel datoteki**: 1) *Inventory movements*, 2) *Sales summary*. "
    "App **sam** prepozna prave zavihke in stolpce."
)

# ----------------- upload -----------------
c1, c2 = st.columns(2)
with c1:
    inv_file = st.file_uploader("Inventory movements (.xlsx/.xls)", type=["xlsx", "xls"], key="inv")
with c2:
    sales_file = st.file_uploader("Sales summary (.xlsx/.xls)", type=["xlsx", "xls"], key="sales")

if not inv_file or not sales_file:
    st.info("Naloži obe datoteki, app nadaljuje samodejno.")
    st.stop()

# ----------------- helperji -----------------
def normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def best_sheet(excel_dict, required_any_sets):
    """
    Izbere sheet, ki ima vse stolpce iz kateregakoli od required_any_sets.
    """
    candidates = []
    for name, df in excel_dict.items():
        cols = {c.strip().lower() for c in df.columns}
        for req in required_any_sets:
            req_norm = {x.lower() for x in req}
            if req_norm.issubset(cols):
                candidates.append(name)
                break
    if candidates:
        # vzamemo tistega z največ vrstic
        return max(candidates, key=lambda n: len(excel_dict[n]))
    return None

def find_column(df: pd.DataFrame, candidates):
    """
    Najde stolpec iz df.columns, ki se ujema z eno od možnih variant.
    """
    cols = list(df.columns)
    lower_map = {c.lower(): c for c in cols}
    strip_map = {c.replace(" ", "").lower(): c for c in cols}
    for cand in candidates:
        cl = cand.lower()
        if cl in lower_map:
            return lower_map[cl]
        if cl.replace(" ", "") in strip_map:
            return strip_map[cl.replace(" ", "")]
    return None

# ----------------- branje datotek -----------------
inv_book = pd.read_excel(inv_file, sheet_name=None)
inv_book = {k: normalize_cols(v) for k, v in inv_book.items()}

sales_book = pd.read_excel(sales_file, sheet_name=None)
sales_book = {k: normalize_cols(v) for k, v in sales_book.items()}

# ----------------- izberi pravi sheet -----------------
inv_required_sets = [
    {"ITEM", "DATE", "INCREASE", "DECREASE"},
    {"ITEM", "DATE", "STOCK QTY"},
    {"Item", "Date", "Increase", "Decrease"},
    {"Item", "Date", "Stock Qty"},
    {"sku", "Date", "Qty In", "Qty Out"},
]

inv_sheet = best_sheet(inv_book, inv_required_sets)
if not inv_sheet:
    st.error(
        "Ne najdem primernega zavihka v Inventory datoteki.\n"
        "Poskrbi, da obstaja zavihek z vsaj stolpci npr. ITEM, DATE in INCREASE/DECREASE ali STOCK QTY."
    )
    st.stop()

sales_required_sets = [
    {"ITEM", "SALES QTY", "SALES VALUE"},
    {"Item", "Sales Qty", "Sales Value"},
    {"sku", "Qty", "Revenue"},
]

sales_sheet = best_sheet(sales_book, sales_required_sets)
if not sales_sheet:
    st.error(
        "Ne najdem primernega zavihka v Sales datoteki.\n"
        "Potrebni stolpci: ITEM, SALES QTY, SALES VALUE."
    )
    st.stop()

inv_df = inv_book[inv_sheet].copy()
sales_df = sales_book[sales_sheet].copy()

# ----------------- mapiranje stolpcev (auto) -----------------
inv_item_col = find_column(inv_df, ["ITEM", "Item", "sku"])
inv_date_col = find_column(inv_df, ["DATE", "Date", "Posting Date", "move_date"])
inv_inc_col = find_column(inv_df, ["INCREASE", "Increase", "Qty In", "QtyIn"])
inv_dec_col = find_column(inv_df, ["DECREASE", "Decrease", "Qty Out", "QtyOut"])
inv_stock_qty_col = find_column(inv_df, ["STOCK QTY", "Stock Qty", "Inventory", "Qty On Hand", "QtyOnHand"])

missing_core = []
if inv_item_col is None:
    missing_core.append("ITEM")
if inv_date_col is None:
    missing_core.append("DATE")

if missing_core:
    st.error(
        f"Inventory zavihek '{inv_sheet}' manjka ključne stolpce: {', '.join(missing_core)}.\n"
        "Prosim preveri, da imaš stolpce za ITEM in DATE."
    )
    st.stop()

sales_item_col = find_column(sales_df, ["ITEM", "Item", "sku"])
sales_qty_col = find_column(sales_df, ["SALES QTY", "Sales Qty", "Qty", "Quantity"])
sales_val_col = find_column(sales_df, ["SALES VALUE", "Sales Value", "Revenue", "Amount"])
sales_desc_col = find_column(sales_df, ["Description", "DESC", "Name", "Item Name"])

missing_sales = []
if sales_item_col is None:
    missing_sales.append("ITEM")
if sales_qty_col is None:
    missing_sales.append("SALES QTY")
if sales_val_col is None:
    missing_sales.append("SALES VALUE")

if missing_sales:
    st.error(
        f"Sales zavihek '{sales_sheet}' manjka ključne stolpce: {', '.join(missing_sales)}.\n"
        "Prosim preveri glave stolpcev v Excelu."
    )
    st.stop()

# ----------------- normalizacija -----------------
inv_df = inv_df.rename(columns={inv_item_col: "ITEM", inv_date_col: "DATE"})
if inv_inc_col:
    inv_df = inv_df.rename(columns={inv_inc_col: "INCREASE"})
else:
    inv_df["INCREASE"] = 0
if inv_dec_col:
    inv_df = inv_df.rename(columns={inv_dec_col: "DECREASE"})
else:
    inv_df["DECREASE"] = 0

if inv_stock_qty_col:
    inv_df = inv_df.rename(columns={inv_stock_qty_col: "STOCK_QTY"})
else:
    # Če nimamo STOCK_QTY, ga izračunamo iz kumulativne razlike INCREASE - DECREASE
    inv_df = inv_df.sort_values(["ITEM", "DATE"])
    inv_df["STOCK_QTY"] = (
        inv_df["INCREASE"].fillna(0) - inv_df["DECREASE"].fillna(0)
    )
    inv_df["STOCK_QTY"] = inv_df.groupby("ITEM")["STOCK_QTY"].cumsum()

inv_df["DATE"] = pd.to_datetime(inv_df["DATE"], errors="coerce")
for c in ["INCREASE", "DECREASE", "STOCK_QTY"]:
    inv_df[c] = pd.to_numeric(inv_df[c], errors="coerce")

sales_df = sales_df.rename(columns={sales_item_col: "ITEM", sales_qty_col: "SALES_QTY", sales_val_col: "SALES_VALUE"})
if sales_desc_col:
    sales_df = sales_df.rename(columns={sales_desc_col: "Description"})
else:
    if "Description" not in sales_df.columns:
        sales_df["Description"] = None

for c in ["SALES_QTY", "SALES_VALUE"]:
    sales_df[c] = pd.to_numeric(sales_df[c], errors="coerce").fillna(0)

as_of = dt.date.today()

# ----------------- izračuni -----------------
# Snapshot: zadnji zapis po datumu za vsak ITEM
inv_df = inv_df.sort_values(["ITEM", "DATE"])
last_idx = inv_df.groupby("ITEM")["DATE"].idxmax()

needed_cols = ["ITEM", "DATE"]
if "STOCK_QTY" in inv_df.columns:
    needed_cols.append("STOCK_QTY")

snapshot = inv_df.loc[last_idx, needed_cols].copy()

snapshot = snapshot.rename(columns={"DATE": "last_move_date"})
if "STOCK_QTY" in snapshot.columns:
    snapshot = snapshot.rename(columns={"STOCK_QTY": "stock_asof"})
else:
    snapshot["stock_asof"] = np.nan

# Zadnji inbound (INCREASE > 0)
inbound = inv_df[inv_df["INCREASE"].fillna(0) > 0].groupby("ITEM")["DATE"].max().rename("last_in_date")
snap = snapshot.merge(inbound, on="ITEM", how="left")

snap["days_since_in"] = (pd.to_datetime(as_of) - snap["last_in_date"]).dt.days
snap["days_since_in"] = snap["days_since_in"].fillna(999999).astype(int)
snap["aging_bucket"] = pd.cut(
    snap["days_since_in"],
    bins=[-1, 30, 60, 90, 120, 10**9],
    labels=["0-30", "31-60", "61-90", "91-120", ">120"],
)

merged = snap.merge(sales_df[["ITEM", "Description", "SALES_QTY", "SALES_VALUE"]], on="ITEM", how="left")

# ----------------- prikaz -----------------
st.subheader("Povzetek prepoznanih podatkov")
k1, k2, k3, k4 = st.columns(4)
k1.metric("Št. artiklov", int(merged["ITEM"].nunique()))
k2.metric("Skupna prodaja (qty)", int(merged["SALES_QTY"].fillna(0).sum()))
k3.metric("Skupna prodaja (value)", f"{merged['SALES_VALUE"].fillna(0).sum():,.2f}")
k4.metric("As-of", str(as_of))

st.caption(f"Inventory sheet: **{inv_sheet}**, Sales sheet: **{sales_sheet}** (samodejno izbrano)")

st.subheader("Snapshot zaloge + prodaja po artiklu")
cols = ["ITEM", "Description", "stock_asof", "last_move_date", "last_in_date", "days_since_in", "aging_bucket", "SALES_QTY", "SALES_VALUE"]
st.dataframe(merged[cols].sort_values("stock_asof", ascending=False), use_container_width=True)

bucket_counts = (
    merged.groupby("aging_bucket")["ITEM"]
    .nunique()
    .reindex(["0-30", "31-60", "61-90", "91-120", ">120"])
    .fillna(0)
    .astype(int)
)
st.markdown("**Razporeditev artiklov po aging bucketih (št. artiklov)**")
st.bar_chart(bucket_counts)

st.download_button("Download merged snapshot (CSV)", merged.to_csv(index=False).encode("utf-8"), "merged_snapshot.csv")

st.success("Samodejni izračun končan.")
