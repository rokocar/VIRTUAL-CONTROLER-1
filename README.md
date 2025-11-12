# Virtual Controller — XLS Minimal (Internet App)
Najbolj preprost način: naložiš **en Excel (.xlsx)** in dobiš 3 poročila (Inventory aging, Reorders, AR aging).

## Hitra namestitev (Streamlit Cloud — brez strežnika)
1. Ustvari nov Git repo z vsebino te mape.
2. Pojdi na **share.streamlit.io** (Streamlit Community Cloud) in poveži repo.
3. Kot glavno datoteko izberi `app.py`. Po deployu dobiš javni URL (https).

## Lokalno
```
pip install -r requirements.txt
streamlit run app.py
```
Odpri http://localhost:8501

## Excel format (zavihki)
- `items`: item_id, sku, name, category, standard_cost, lead_time_days, supplier_id
- `inventory_balances`: item_id, location, qty_on_hand, unit_cost, as_of_date
- `stock_moves`: item_id, direction(in/out), qty, move_date
- `purchase_orders`: item_id, qty_ordered, qty_received, expected_receipt_date
- `invoices_ar`: invoice_id, customer_id, due_date, open_amount
- `customers`: customer_id, name

> Dovolj je minimalni nabor stolpcev; manjkajoče polja bo app zapolnil z ničlami oz. privzetimi vrednostmi (npr. lead_time=14).
