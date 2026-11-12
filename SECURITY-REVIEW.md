# Security Review

Date: 2026-11-12

This read-only review identified the following issues. No application code was changed as part of this review.

| # | Severity | File | Lines | Vulnerability | Confidence |
|---|----------|------|-------|---------------|------------|
| 1 | 🟡 MEDIUM | `Backend/Price_Doc/pricesched.py` | 379, 413-430 | Uploaded document values can become Excel formulas when they begin with `=`, `+`, `-`, or `@`, enabling formula injection when the workbook is opened. | 9/10 |
| 2 | 🟡 MEDIUM | `Backend/tender_intake/app.py` | 163-173 | Download authorization relies only on the user-supplied job ID; anyone who obtains another job ID may download its generated documents. | 8/10 |

## Recommended remediation

1. Sanitize extracted workbook values before writing them to Excel, or explicitly write untrusted values as strings.
2. Bind generated jobs to an authenticated user or creating session and verify ownership before serving downloads.
