
import argparse
import os

from excel_writer import build_workbook
from extractor import extract_report


def process_files(paths, out_path):
    reports = []
    for p in paths:
        rep = extract_report(p)
        rep["source"] = os.path.basename(p)
        reports.append(rep)
    build_workbook(reports, out_path)
    return reports


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Extract SSET report(s) into 5 editable Excel tables")
    ap.add_argument("files", nargs="+")
    ap.add_argument("-o", "--out", default="SSET_tables.xlsx")
    a = ap.parse_args()
    reps = process_files(a.files, a.out)
    for r in reps:
        print(f"{r['source']}: {len(r['fields'])} fields, {len(r['bidders'])} bidder(s)")
    print("Saved ->", a.out)
