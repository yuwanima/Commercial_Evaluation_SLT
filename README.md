# Commercial_Evaluation_SLT


## Run [app.py](Backend/tender_intake/app.py)
From a PowerShell window:

```powershell
cd Backend/tender_intake
pip install -r ../requirements.txt
$env:SECRET_KEY = python -c "import secrets; print(secrets.token_urlsafe(32))"
$env:DATABASE_URL = "sqlite:///tender_local.db"
python app.py
```

Then open http://127.0.0.1:5000. The local secret is generated for that run; set a stable secret in your deployment environment for production use. Do not put production secrets into the source ZIP or commit them to version control.
In the browser, choose **Process all documents** and upload the technical evaluation report, bidder price schedules, approved tender strategy/evaluation weighting record, and CET evaluation report. Add a Tender Board/authorized approval resolution or formal award notice to confirm the actual awarded bidder.

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


## Post-technical commercial evaluation and award assessment

In **Process all documents** mode, the system now builds an **Award Assessment** after the TET report and price schedules have been extracted. It:

1. Uses the TET report's overall bidder result and explicitly identified mandatory/drop-dead criteria to establish who may proceed to commercial evaluation. A non-mandatory non-compliance is recorded for review; it is not treated as an automatic disqualification when the TET report explicitly records the bidder as qualified. Ambiguous or conflicting TET outcomes are held for manual review.
2. Extracts bidder quoted totals from the price schedules. When no filled-in price schedule is parsed for a bidder, an explicit bidder/quoted-total table in the TET report may be used as a clearly labelled fallback.
3. Compares eligible bidders only when the prices can be compared in the same stated currency. It does not convert currencies or add a separate local-agent commission whose currency is unclear.
4. Reads the technical and commercial weighting percentages from an uploaded approval criteria PDF and stores them with bidder scores. Technical scores come from the TET report; commercial scores are calculated from comparable price schedules as `(lowest evaluated price / bidder evaluated price) × 100`.
5. Calculates weighted technical and commercial scores as `score × approved weighting / 100`, then sums them for the overall score. A weighted recommendation requires a score for each technically eligible bidder, one comparable currency, and approved weightings totaling 100%. Otherwise it shows the lowest eligible price as a **provisional price leader**, or reports that the available evidence is insufficient.

The award workbook contains the bidder-by-bidder evidence, price ranks, scores, data sources and outstanding checks. Download it using **Award Assessment Excel**.

The screen separates a **preferred bidder / provisional price leader** from an **actually awarded bidder**. It only displays an actually awarded bidder when an uploaded Tender Board/authority approval record or formal award notice explicitly identifies that bidder and the bidder matches a technically eligible bidder in the TET records. A TET/CET recommendation alone is not treated as proof of an award.

### Documents that produce a complete weighted recommendation

Upload the TET report, bidder price schedules, and the approved criteria PDF showing technical and commercial weightages. The TET report must contain a clear technical qualification outcome and numeric technical scores if a composite ranking is required. To confirm an actual award, also upload the Tender Board/authorized approval resolution or formal letter/notice of award.

**Governance safeguard:** the system produces an evaluation result or recommendation for review; it never itself awards the tender. The preferred bidder remains subject to TET/CET review, documented clarifications/negotiations where applicable, the relevant Tender Board/DOA approval, and contract execution. A price-only leader is not labelled as an awarded bidder.
