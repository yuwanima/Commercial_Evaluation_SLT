
import argparse
import os

import excel_writer
import extractor
import tet_extractor
import tet_writer

REPORT_TYPES = {
    "sset": {
        "tab": "SSET",
        "title": "SSET - Sole Source Evaluation Report",
        "help": "Upload one or more SSET evaluation reports (.docx).",
        "extract": extractor.extract_report,
        "build": excel_writer.build_workbook,
        "colour": "#1F3864",
    },
    "tet": {
        "tab": "TET",
        "title": "TET - Technical Evaluation Report",
        "help": "Upload one or more TET evaluation reports (.docx). Tables 3-5 get one column per bidder.",
        "extract": tet_extractor.extract_tet,
        "build": tet_writer.build_workbook,
        "colour": "#0B6E4F",
    },
}


def process_files(kind, paths, out_path):
    spec = REPORT_TYPES[kind]
    reports = []
    for p in paths:
        rep = spec["extract"](p)
        rep["source"] = os.path.basename(p)
        reports.append(rep)
    spec["build"](reports, out_path)
    return reports


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Extract evaluation report(s) into 5 editable Excel tables")
    ap.add_argument("--type", choices=list(REPORT_TYPES), default="sset")
    ap.add_argument("files", nargs="+")
    ap.add_argument("-o", "--out", default="tables.xlsx")
    a = ap.parse_args()
    for r in process_files(a.type, a.files, a.out):
        print(f"{r['source']}: {len(r['fields'])} fields, {len(r['bidders'])} bidder(s)")
    print("Saved ->", a.out)
