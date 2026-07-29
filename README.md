# CLAIMS-RECONCILIATION-ENGINe

## Overview

The **Claims Reconciliation Engine** is a Python-based automation tool developed to reconcile insurance claims payable transactions stored in Microsoft Excel workbooks.

The application automatically matches debit and credit transactions using multiple reconciliation strategies and produces a detailed reconciliation report with audit information for every matched transaction.

The objective is to reduce manual reconciliation effort, improve matching accuracy, and provide a transparent audit trail for finance teams.

---

# Features

* Automated reconciliation of claims payable transactions
* Excel file selection through a graphical file picker
* Automatic validation of required columns
* Cleans and standardizes financial data
* Removes summary rows such as Closing Balance, Group Total and Grand Total
* Extracts claim numbers using Regular Expressions
* Builds optimized lookup tables for fast reconciliation
* Supports multiple reconciliation methods
* Generates reconciliation statistics
* Exports multiple Excel reports automatically

---

# Reconciliation Logic

The reconciliation engine performs matching in the following order.

## Stage 1 — CNI → DNI Matching

Matches:

* Claim Number
* Transaction Type
* Amount

This is considered the highest confidence match.

---

## Stage 2 — CNI → PVB Matching

If no DNI transaction is found, the engine searches PVB debit transactions using:

* Claim Number
* Amount

---

## Stage 3 — Amount Only Matching

If no claim-based match is found, the engine performs a fallback reconciliation using:

* Exact Amount

Transactions reconciled using this method are marked as:

```
AMOUNT ONLY
```

allowing them to be reviewed separately if required.

---

# Data Cleaning

Before reconciliation begins, the application automatically:

* Converts Debit values to numeric
* Converts Credit values to numeric
* Rounds financial values to two decimal places
* Removes blank transactions
* Removes total rows
* Resets DataFrame indexes

This ensures consistent processing regardless of the input file format.

---

# Lookup Tables

To improve performance, lookup dictionaries are created for:

* Debit Amounts
* DNI Claim Numbers
* PVB Claim Numbers

Using lookup tables significantly reduces processing time compared to searching the entire dataset repeatedly.

---

# Output

The application automatically creates a new Excel workbook containing:

## 1. Reconciled

All successfully matched transactions.

---

## 2. Unreconciled

Transactions that could not be matched.

---

## 3. All Transactions

Original transaction list including reconciliation status and audit columns.

Additional columns include:

* Status
* Match Type
* Matched Row
* Matched Transaction
* Matched Document
* Matched Date

---

## 4. Summary

Overall reconciliation statistics including:

* Total Transactions
* Total Debits
* Total Credits
* Reconciled Transactions
* Unreconciled Transactions
* Matched Debits
* Matched Credits
* Outstanding Debits
* Outstanding Credits
* Outstanding Balance

---

## 5. Match Analysis

Breakdown showing how each transaction was reconciled.

Example:

| Match Type  | Count |
| ----------- | ----: |
| CNI-DNI     |   214 |
| CNI-PVB     |    89 |
| AMOUNT ONLY |    37 |

---

# Technologies Used

* Python
* Pandas
* OpenPyXL
* Tkinter
* Regular Expressions (re)
* Collections (defaultdict)
* Microsoft Excel

---

# Project Workflow

```
Select Excel File
        │
        ▼
Load Workbook
        │
        ▼
Validate Required Columns
        │
        ▼
Clean Transaction Data
        │
        ▼
Extract Claim Numbers
        │
        ▼
Build Lookup Tables
        │
        ▼
CNI → DNI Matching
        │
        ▼
CNI → PVB Matching
        │
        ▼
Amount Matching
        │
        ▼
Generate Reports
        │
        ▼
Export Excel Workbook
```


# Example Console Output

======================================================================
CLAIMS RECONCILIATION ENGINE
======================================================================

BUILDING DEBIT LOOKUPS

DNI Claims Indexed : 1,245
PVB Claims Indexed : 328
Amount Records     : 2,417

STARTING RECONCILIATION

Reconciliation completed.

======================================================================
RECONCILIATION SUMMARY
======================================================================

Total Transactions       : 7,842
Reconciled Transactions  : 7,603
Unreconciled Transactions: 239

Outstanding Balance      : 12,540.35
```

---

# Installation

Clone the repository

```bash
git clone https://github.com/YOUR_USERNAME/claims-reconciliation-engine.git
```

Install the required packages

```bash
pip install -r requirements.txt
```

---

# Requirements

```
Python 3.10+
pandas
openpyxl
```

Install packages using

```bash
pip install pandas openpyxl
```

---

# Running the Application

Run the script

```bash
python claims_reconciliation.py
```

A file selection window will appear.

Select the Claims Payable Excel workbook and the reconciliation process will begin automatically.

The reconciled workbook will be saved in the same folder as the source file with the suffix:

```
_Reconciled.xlsx
```

---

# Project Structure

```
claims-reconciliation-engine/

│── claims_reconciliation.py
│── README.md
│── requirements.txt
│── LICENSE
│── sample_data/
│── screenshots/
```

---

# Future Improvements

Planned enhancements include:

* Weighted confidence scoring
* Fuzzy narration matching
* Partial payment reconciliation
* One-to-many and many-to-one reconciliation
* SQL database integration
* Streamlit web dashboard
* PDF reconciliation reports
* Logging and exception handling
* Unit tests
* Docker support

---

# Skills Demonstrated

This project demonstrates practical experience with:

* Financial Data Analysis
* Insurance Claims Processing
* Data Cleaning
* Python Automation
* Excel Automation
* Regular Expressions
* Lookup Optimization
* Algorithm Design
* Financial Reconciliation
* Report Generation

---

# License

This project is licensed under the MIT License.

---


