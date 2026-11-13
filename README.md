# Commercial_Evaluation_SLT


## Run[app.py](Backend/tender_intake/app.py)
    pip install -r requirements.txt
    python app.py --open-local        # open http://127.0.0.1:5000
 - Choose one or more SSET .docx reports -> **Process** -> **View**.
 - `--open-local`: View opens the workbook straight in Excel (run the app on the member's PC).
   Without it, View downloads the .xlsx.

## Command line
    Web front-end for the GCC / SCC bidder compliance consolidator.

    pip install flask openpyxl python-docx pdfplumber
    python app.py            ->  open http://127.0.0.1:5000

## Output workbook (one detail tab per company)
    Bid compliance consolidator  (GCC Excel/PDF + SCC Word/PDF  ->  comparison Excel files)

    Usage
    -----
        python consolidate.py  <files / folders / .zip> ...  [-o OUTPUT_DIR]

    Give it the bidders' returned files (any mix, two or more bidders). It will
    * read every GCC compliance sheet (.xlsx / .pdf)  -> GCC_Compliance_Comparison.xlsx
    * read every SCC form (.docx / .doc / .pdf)       -> SCC_Compliance_Comparison.xlsx
    Each bidder gets its own column group so all bidders can be compared clause by clause.
    Explanatory notes are wrapped and displayed as separate points in the Excel cells for easier reading.

    The web front-end also has a **Price Schedule** tab. Upload completed local or
    foreign schedules to view calculated totals and currency-aware rankings, then
    download the detailed `Price_Schedules_<TenderNo>.xlsx` workbook. The workbook
    uses the Annexure 01 price-evaluation layout, with one separately populated
    worksheet per company. Each tab contains the bidder, item, quantity, price,
    tax, ranking, and source columns from that company's input document.

    Bidder name
    -----------
    * file inside a sub-folder  -> the sub-folder name   (Bidders/ABC Ltd/GCC.xlsx -> "ABC Ltd")
    * file given directly       -> file name with GCC/SCC/etc. words removed
    * explicit                  -> "ABC Ltd=path/to/file.xlsx"
