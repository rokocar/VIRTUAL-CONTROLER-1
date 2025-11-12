# Virtual Controller — Auto (Two Files)
Brez nastavitev: naloži **dve** Excel datoteki (inventory movements + sales summary). App sam prepozna prave zavihke in stolpce ter izračuna:
- snapshot zaloge (as-of danes),
- preprost aging (zadnji inbound),
- združitev s prodajo, KPI-je, graf in CSV export.

## Zagon
```
pip install -r requirements.txt
streamlit run app.py
```
