# Commercial_Evaluation_SLT


# SSET Report -> Editable Excel Tables

## Run
    pip install -r requirements.txt
    python app.py --open-local        # open http://127.0.0.1:5000
 - Choose one or more SSET .docx reports -> **Process** -> **View**.
 - `--open-local`: View opens the workbook straight in Excel (run the app on the member's PC).
   Without it, View downloads the .xlsx.

## Command line
    python process.py report1.docx report2.docx -o tables.xlsx

## Output workbook (one sheet per table)
 01 Tender & Timeline | 02 Background Summary | 03 Bidder Details |
 04 Compliance & Observations | 05 Commercial Evaluation | Extraction Log
 - One column per report (tables 1,2,4,5) or per bidder (table 3).
 - Yellow cell = nothing found, fill in manually. Dropdowns on compliance / Yes-No rows.
 - Extraction Log lists missing fields and warnings per report.

