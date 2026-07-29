# CLAIMS-RECONCILIATION-ENGINE
## Overview

The Claims Reconciliation Engine is a Python-based financial reconciliation system developed to automate the reconciliation of insurance claims payable transactions.

The application analyzes debit and credit transactions from Excel files and automatically identifies matching transactions using multiple reconciliation strategies.

The goal is to reduce manual reconciliation effort while improving accuracy and providing a complete audit trail.

---

## Features

✔ Exact amount reconciliation

✔ Claim number matching

✔ Document number matching

✔ Reference matching

✔ Split payment reconciliation

✔ Multiple credits → one debit

✔ One credit → multiple debits

✔ Claim-based combination matching

✔ Fuzzy narration matching

✔ Automatic Excel report generation

✔ Reconciliation summary statistics

✔ Audit trail for every matched transaction

---

## Technologies

- Python
- Pandas
- OpenPyXL
- Regular Expressions
- Difflib
- Tkinter

---

## Reconciliation Workflow

1. Import Claims Payable Excel
2. Clean transaction data
3. Build lookup tables
4. Execute reconciliation passes
5. Generate reconciliation report
6. Export results to Excel

---

## Matching Algorithms

### Pass 1
Exact Amount Match

### Pass 2
Claim Number Match

### Pass 3
Document Number Match

### Pass 4
Reference Match

### Pass 5
Multiple Credits → One Debit

### Pass 6
One Credit → Multiple Debits

### Pass 7
Claim Combination Match

### Pass 8
Fuzzy Narration Match

---

## Output

The application produces:

- Reconciled transactions
- Outstanding transactions
- Complete transaction listing
- Summary statistics

---

## Example Output

| Metric | Value |
|---------|------:|
| Total Transactions | 7,582 |
| Reconciled | 7,406 |
| Outstanding | 176 |
| Reconciliation Rate | 97.68% |

---

## Future Improvements

- Weighted scoring engine
- Machine learning-assisted reconciliation
- SQL database integration
- Streamlit dashboard
- REST API
- PDF reconciliation reports
- Unit testing
- Docker support

---

## Author

Wendy Awuor
