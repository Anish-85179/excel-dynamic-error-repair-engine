# Enterprise Financial Formula Auditor & Dynamic Error Repair Engine (Block 2)

An enterprise-grade financial formula auditing, automated error-trapping, and subledger reconciliation pipeline built in Microsoft Excel. This project establishes robust data hygiene frameworks to sanitize corrupted general ledger records, resolve formula calculation errors (`#N/A`, `#DIV/0!`, `#VALUE!`), execute precise variance reconciliation against subledger balances, and drive executive visibility through a real-time KPI console.

---

## 🏗️ Architectural Overview & Structure

The workbook is structured into **4 sequential operational sheets** designed to ingest raw messy accounting data, sanitize financial calculations, verify ledger-to-subledger alignment, and present executive metrics:

1. **`01_Corrupted_General_Ledger`** — Ingestion layer containing 100 raw transaction records with deliberate enterprise anomalies, text-formatting corruptions, invalid vendor IDs, and broken calculations.
2. **`02_Formula_Repair_Engine`** — Transformation layer utilizing advanced Excel error-trapping and cleaning formulas (`VALUE`, `TRIM`, `SUBSTITUTE`, `IFORMAT`) to normalize vendor codes, parse quantities, and calculate clean extended costs.
3. **`03_Reconciliation_Matrix`** — Variance detection engine comparing repaired extended costs against target subledger balances using a strict **$0.05 tolerance threshold**, automatically flagging matched records versus required manual reviews.
4. **`04_Audit_Executive_Console`** — Executive KPI dashboard tracking total audits, trapped errors, reconciliation variances, and system health stability scores in real time.

---

## 🛠️ Core Technical Implementation & Formulas

### 1. Data Cleaning & Error-Trapping (`02_Formula_Repair_Engine`)
To resolve text-string price formats and formula errors (`#DIV/0!`, `#VALUE!`), desktop-compatible sanitization formulas are applied:
* **Clean Price Parsing:**
  ```excel
  =IFERROR(VALUE(TRIM(SUBSTITUTE(SUBSTITUTE('01_Corrupted_General_Ledger'!F2, "$", ""), " ", ""))), 0)
