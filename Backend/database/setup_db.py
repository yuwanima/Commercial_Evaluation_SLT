"""One-time setup. Run on the server after setting DATABASE_URL:
    python setup_db.py  path/to/IF_Buying_Selling_Exchange_Rates.xlsx
Creates the tables and loads the CBSL USD rate history. Safe to run again (rates are updated, not duplicated)."""
import sys, tender_data as t

if len(sys.argv) != 2:
    sys.exit("usage: python setup_db.py IF_Buying_Selling_Exchange_Rates.xlsx")
t.init()
rep = t.load_cbsl_xlsx(sys.argv[1])
print(f"loaded {rep['loaded']} daily rates, {rep['first']} to {rep['last']}")
for d, why in rep["skipped"]:
    print("NOT loaded (check against CBSL):", d, "-", why)
