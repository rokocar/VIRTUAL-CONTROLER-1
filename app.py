import streamlit as st
import pandas as pd
import datetime as dt

st.set_page_config(page_title="Virtual Controller (XLS Minimal)", layout="wide")
st.title("Virtual Controller — XLS Minimal")

with st.sidebar:
    st.markdown("### 1) Naloži Excel")
    uploaded = st.file_uploader("Excel (.xlsx) z zavihki: items, inventory_balances, stock_moves, purchase_orders, invoices_ar, customers", type=["xlsx"])
    as_of = st.date_input("As-of date", value=dt.date.today())
    horizon_days = st.number_input("Horizon (reorders)", 1, 60, value=7)
    z_score = st.number_input("Service level Z", 0.0, 5.0, value=1.65)
    demand_window_days = st.number_input("Demand window (days)", 7, 365, value=60)

def must_have(df_dict, names):
    missing = [n for n in names if n not in df_dict]
    if missing:
        st.error(f"Manjkajo zavihki v Excelu: {', '.join(missing)}")
        st.stop()

@st.cache_data(show_spinner=False)
def read_excel(file):
    try:
        x = pd.read_excel(file, sheet_name=None)
        return x
    except Exception as e:
        st.error(f"Napaka pri branju Excela: {e}")
        st.stop()

def kpi_currency(v): 
    return f"{v:,.2f}"

if uploaded:
    data = read_excel(uploaded)
    must_have(data, ["items","inventory_balances","stock_moves","purchase_orders","invoices_ar","customers"])

    # Normalize minimal required columns
    items = data["items"].copy()
    invb  = data["inventory_balances"].copy()
    moves = data["stock_moves"].copy()
    pos   = data["purchase_orders"].copy()
    ar    = data["invoices_ar"].copy()
    cust  = data["customers"].copy()

    # Ensure expected columns exist (with fallbacks where possible)
    for col in ["item_id","sku","name","category","standard_cost","lead_time_days","supplier_id"]:
        if col not in items.columns:
            items[col] = None
    for col in ["item_id","location","qty_on_hand","unit_cost","as_of_date"]:
        if col not in invb.columns:
            invb[col] = None
    for col in ["item_id","direction","qty","move_date"]:
        if col not in moves.columns:
            moves[col] = None
    for col in ["item_id","qty_ordered","qty_received","expected_receipt_date"]:
        if col not in pos.columns:
            pos[col] = None
    for col in ["invoice_id","customer_id","due_date","open_amount"]:
        if col not in ar.columns:
            ar[col] = None
    for col in ["customer_id","name"]:
        if col not in cust.columns:
            cust[col] = None

    # Coerce dtypes
    invb["as_of_date"] = pd.to_datetime(invb["as_of_date"], errors="coerce")
    moves["move_date"] = pd.to_datetime(moves["move_date"], errors="coerce")
    pos["expected_receipt_date"] = pd.to_datetime(pos["expected_receipt_date"], errors="coerce")
    ar["due_date"] = pd.to_datetime(ar["due_date"], errors="coerce")

    # ---------------- Inventory aging ----------------
    st.header("Inventory aging")
    inbound = moves.loc[moves["direction"].str.lower().eq("in", na=False)].groupby("item_id")["move_date"].max().rename("last_in_date")
    df = invb.merge(items, on="item_id", how="left").merge(inbound, on="item_id", how="left")
    df = df[df["as_of_date"].dt.date == as_of]
    unit_cost = df["unit_cost"].fillna(df["standard_cost"]).fillna(0)
    df["value"] = (df["qty_on_hand"].fillna(0) * unit_cost).astype(float)
    days = (pd.to_datetime(as_of) - df["last_in_date"]).dt.days.fillna(999999).astype(int)
    df["aging_bucket"] = pd.cut(days, bins=[-1,30,60,90,120,10**9], labels=["0-30","31-60","61-90","91-120",">120"])

    kpi_cols = st.columns(3)
    kpi_cols[0].metric("Total Qty", int(df["qty_on_hand"].fillna(0).sum()))
    kpi_cols[1].metric("Total Value", kpi_currency(df["value"].sum()))
    kpi_cols[2].metric(">120d Value", kpi_currency(df.loc[df["aging_bucket"].astype(str)==">120","value"].sum()))

    st.dataframe(df[["sku","name","category","location","qty_on_hand","unit_cost","value","aging_bucket"]].round(2), use_container_width=True)

    # ---------------- Reorders (ROP) ----------------
    st.header("Reorder suggestions")
    cutoff = pd.to_datetime(as_of) - pd.Timedelta(days=int(demand_window_days))
    out_moves = moves[(moves["direction"].str.lower()=="out") & (moves["move_date"]>=cutoff)]
    demand = out_moves.groupby("item_id")["qty"].sum() / max(int(demand_window_days),1)
    sigma = out_moves.groupby("item_id")["qty"].std().fillna(0)
    lead = items.set_index("item_id")["lead_time_days"].fillna(14)

    onhand = invb.groupby("item_id")["qty_on_hand"].sum()
    onorder = pos.assign(open_qty=lambda d: (d["qty_ordered"].fillna(0)-d["qty_received"].fillna(0))).groupby("item_id")["open_qty"].sum()

    idx = items["item_id"]
    dpp = demand.reindex(idx).fillna(0)
    sdp = sigma.reindex(idx).fillna(0)
    ld = lead.reindex(idx).fillna(14)
    oh = onhand.reindex(idx).fillna(0)
    oo = onorder.reindex(idx).fillna(0)

    safety_stock = z_score * sdp * (ld**0.5)
    reorder_point = dpp*ld + safety_stock
    suggested = ((dpp*(ld+horizon_days)) + safety_stock - oh - oo).clip(lower=0).round()

    reo = items.copy()
    reo["demand_per_day"] = dpp.round(2)
    reo["lead_time_days"] = ld
    reo["safety_stock"] = safety_stock.round(1)
    reo["reorder_point"] = reorder_point.round(1)
    reo["on_hand"] = oh
    reo["on_order"] = oo
    reo["suggested_qty"] = suggested.astype(int)

    st.dataframe(reo[["sku","name","supplier_id","demand_per_day","lead_time_days","safety_stock","reorder_point","on_hand","on_order","suggested_qty"]], use_container_width=True)

    # ---------------- AR aging ----------------
    st.header("AR aging")
    ar2 = ar.copy()
    ar2["days_past_due"] = (pd.to_datetime(as_of) - ar2["due_date"]).dt.days.clip(lower=0)
    bins = [0,30,60,90,120,10**9]
    labels = ["b0_30","b31_60","b61_90","b91_120","b121_plus"]
    ar2["bucket"] = pd.cut(ar2["days_past_due"], bins=bins, labels=labels, right=True, include_lowest=True)

    pivot = ar2.pivot_table(index="customer_id", columns="bucket", values="open_amount", aggfunc="sum", fill_value=0).reset_index()
    pivot = pivot.merge(cust[["customer_id","name"]], on="customer_id", how="left")
    pivot["total_open"] = pivot[labels].sum(axis=1)
    k = {
        "Total open": pivot["total_open"].sum(),
        ">120": pivot["b121_plus"].sum()
    }
    kpi2 = st.columns(2)
    kpi2[0].metric("Total open", kpi_currency(k["Total open"]))
    kpi2[1].metric(">120", kpi_currency(k[">120"]))

    st.dataframe(pivot[["customer_id","name"]+labels+["total_open"]].round(2), use_container_width=True)

    st.success("Ready. Vse izračuni tečejo **lokalno nad tvojim Excelom**.")

else:
    st.info("⬆️ Naloži Excel datoteko, nato vidiš Inventory aging, Reorders in AR aging.")
