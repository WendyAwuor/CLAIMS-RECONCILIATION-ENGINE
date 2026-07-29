import pandas as pd
from collections import defaultdict
from tkinter import Tk, filedialog
import os
import re


# =====================================================
# SELECT INPUT FILE
# =====================================================

Tk().withdraw()

file_path = filedialog.askopenfilename(
    title="Select Claims Payable Excel File",
    filetypes=[
        ("Excel Files", "*.xlsx *.xls")
    ]
)

if not file_path:
    raise SystemExit("No file selected.")


print("=" * 70)
print("CLAIMS RECONCILIATION ENGINE")
print("=" * 70)

print(f"Input File: {file_path}")


# =====================================================
# LOAD DATA
# =====================================================

df = pd.read_excel(
    file_path,
    header=16
)


# =====================================================
# REQUIRED COLUMNS CHECK
# =====================================================

required_columns = [
    "Document Date",
    "Document Number",
    "Transaction Code",
    "Reference",
    "Narration",
    "Debits",
    "Credits"
]


missing = [
    col for col in required_columns
    if col not in df.columns
]


if missing:
    raise Exception(
        f"Missing columns: {missing}"
    )

# =====================================================
# CLEAN NUMERIC VALUES
# =====================================================

df["Debits"] = (
    pd.to_numeric(
        df["Debits"],
        errors="coerce"
    )
    .fillna(0)
    .round(2)
)


df["Credits"] = (
    pd.to_numeric(
        df["Credits"],
        errors="coerce"
    )
    .fillna(0)
    .round(2)
)

# =====================================================
# REMOVE TOTAL ROWS
# =====================================================

df = df[
    ~df["Narration"]
    .astype(str)
    .str.contains(
        r"Closing\s*Balance|Group\s*Total|Grand\s*Total",
        case=False,
        regex=True,
        na=False
    )
].copy()

# =====================================================
# KEEP TRANSACTION ROWS ONLY
# =====================================================

df = df[
    (df["Debits"] != 0)
    |
    (df["Credits"] != 0)
].copy()


df.reset_index(
    drop=True,
    inplace=True
)

# =====================================================
# RESULT COLUMNS
# =====================================================

df["Status"] = "Unreconciled"

df["Match Type"] = ""

df["Matched Row"] = pd.Series(
    "",
    index=df.index,
    dtype="object"
)

df["Matched Transaction"] = ""

df["Matched Document"] = ""

df["Matched Date"] = pd.NaT

print()
print("DATA SUMMARY")
print("-" * 50)

print(
    f"Rows Loaded : {len(df):,}"
)

print(
    f"Total Debits : {df['Debits'].sum():,.2f}"
)

print(
    f"Total Credits: {df['Credits'].sum():,.2f}"
)

print()

# =====================================================
# CLAIM EXTRACTION
# =====================================================

claim_pattern = re.compile(
    r"\d{2}/\d{2}/[A-Z]+/\d+",
    re.IGNORECASE
)

def extract_claim_numbers(text):

    if pd.isna(text):
        return []

    return [
        x.upper()
        for x in claim_pattern.findall(
            str(text)
        )
    ]

# =====================================================
# TRANSACTION TYPE
# =====================================================

def get_transaction_code(row):

    return (
        str(row["Transaction Code"])
        .upper()
        .strip()
    )

# =====================================================
# RECONCILIATION UPDATE FUNCTION
# =====================================================

def reconcile(
    credit_idx,
    debit_idx,
    match_type
):

    df.at[
        credit_idx,
        "Status"
    ] = "Reconciled"


    df.at[
        debit_idx,
        "Status"
    ] = "Reconciled"


    df.at[
        credit_idx,
        "Match Type"
    ] = match_type


    df.at[
        debit_idx,
        "Match Type"
    ] = match_type



    df.at[
        credit_idx,
        "Matched Row"
    ] = str(debit_idx + 2)


    df.at[
        debit_idx,
        "Matched Row"
    ] = str(credit_idx + 2)



    df.at[
        credit_idx,
        "Matched Transaction"
    ] = df.at[
        debit_idx,
        "Transaction Code"
    ]


    df.at[
        debit_idx,
        "Matched Transaction"
    ] = df.at[
        credit_idx,
        "Transaction Code"
    ]



    df.at[
        credit_idx,
        "Matched Document"
    ] = str(
        df.at[
            debit_idx,
            "Document Number"
        ]
    )


    df.at[
        debit_idx,
        "Matched Document"
    ] = str(
        df.at[
            credit_idx,
            "Document Number"
        ]
    )



    df.at[
        credit_idx,
        "Matched Date"
    ] = df.at[
        debit_idx,
        "Document Date"
    ]


    df.at[
        debit_idx,
        "Matched Date"
    ] = df.at[
        credit_idx,
        "Document Date"
    ]

# =====================================================
# BUILD DEBIT LOOKUP TABLES
# =====================================================

dni_lookup = defaultdict(list)

pvb_lookup = defaultdict(list)

amount_lookup = defaultdict(list)

# =====================================================
# INDEX ALL DEBIT TRANSACTIONS
# =====================================================

debit_rows = df[
    df["Debits"] > 0
].copy()

print("=" * 70)
print("BUILDING DEBIT LOOKUPS")
print("=" * 70)

for idx, row in debit_rows.iterrows():

    tx_code = get_transaction_code(row)


    # -------------------------------------------------
    # AMOUNT LOOKUP
    # -------------------------------------------------

    amount_lookup[
        row["Debits"]
    ].append(idx)


    # -------------------------------------------------
    # DNI TRANSACTIONS
    #
    # Example:
    #
    # DNI100
    # Reference:
    # 23/10/HOD/0679
    #
    # -------------------------------------------------

    if tx_code.startswith("DNI"):


        claims = extract_claim_numbers(
            row["Reference"]
        )


        for claim in claims:

            dni_lookup[
                claim
            ].append(idx)



    # -------------------------------------------------
    # PVB TRANSACTIONS
    #
    # Example:
    #
    # PVB100
    # Narration:
    # CLAIM PAYMENT
    # 23/10/HOD/0679
    #
    # -------------------------------------------------

    elif tx_code.startswith("PVB"):


        claims = extract_claim_numbers(
            row["Narration"]
        )


        for claim in claims:

            pvb_lookup[
                claim
            ].append(idx)

print(
    f"DNI Claims Indexed : {len(dni_lookup):,}"
)

print(
    f"PVB Claims Indexed : {len(pvb_lookup):,}"
)

print(
    f"Amount Records     : {len(amount_lookup):,}"
)

print("=" * 70)
print()
# =====================================================
# RECONCILIATION ENGINE
# =====================================================

print("=" * 70)
print("STARTING RECONCILIATION")
print("=" * 70)

# =====================================================
# PROCESS CREDIT TRANSACTIONS ONLY
# =====================================================

for credit_idx, credit_row in df[
    df["Credits"] > 0
].iterrows():


    # Skip already reconciled rows

    if df.at[
        credit_idx,
        "Status"
    ] == "Reconciled":

        continue



    tx_code = get_transaction_code(
        credit_row
    )


    credit_amount = credit_row[
        "Credits"
    ]


    matched = False



    # =================================================
    # ONLY PROCESS CNI AS CLAIM SOURCE
    # =================================================

    if tx_code.startswith("CNI"):


        claims = extract_claim_numbers(
            credit_row["Reference"]
        )


        # =================================================
        # STAGE 1
        # CNI --> DNI
        # =================================================


        for claim in claims:


            possible_dni = dni_lookup.get(
                claim,
                []
            )


            for debit_idx in possible_dni:


                debit_row = df.loc[
                    debit_idx
                ]


                # Ignore used debit rows

                if debit_row["Status"] == "Reconciled":

                    continue



                debit_amount = debit_row[
                    "Debits"
                ]



                if abs(
                    debit_amount -
                    credit_amount
                ) <= 0.01:


                    reconcile(
                        credit_idx,
                        debit_idx,
                        "CNI-DNI"
                    )


                    matched = True

                    break



            if matched:

                break



        # =================================================
        # STAGE 2
        # CNI --> PVB
        # =================================================


        if not matched:


            for claim in claims:


                possible_pvb = pvb_lookup.get(
                    claim,
                    []
                )



                for debit_idx in possible_pvb:


                    debit_row = df.loc[
                        debit_idx
                    ]



                    if debit_row["Status"] == "Reconciled":

                        continue



                    debit_amount = debit_row[
                        "Debits"
                    ]



                    if abs(
                        debit_amount -
                        credit_amount
                    ) <= 0.01:



                        reconcile(
                            credit_idx,
                            debit_idx,
                            "CNI-PVB"
                        )


                        matched = True

                        break



                if matched:

                    break

    # =================================================
    # STAGE 3
    # AMOUNT ONLY FALLBACK
    # =================================================
    if not matched:


        possible_amounts = amount_lookup.get(
            credit_amount,
            []
        )



        for debit_idx in possible_amounts:


            debit_row = df.loc[
                debit_idx
            ]



            if debit_row["Status"] == "Reconciled":

                continue



            reconcile(
                credit_idx,
                debit_idx,
                "AMOUNT ONLY"
            )


            matched = True

            break

print("Reconciliation completed.")

print()
# =====================================================
# RESULTS
# =====================================================
reconciled = df[
    df["Status"] == "Reconciled"
].copy()

unreconciled = df[
    df["Status"] == "Unreconciled"
].copy()

# =====================================================
# SUMMARY CALCULATIONS
# =====================================================
total_debits = df[
    "Debits"
].sum()

total_credits = df[
    "Credits"
].sum()

matched_debits = reconciled[
    "Debits"
].sum()

matched_credits = reconciled[
    "Credits"
].sum()

unmatched_debits = unreconciled[
    "Debits"
].sum()

unmatched_credits = unreconciled[
    "Credits"
].sum()

difference = (
    unmatched_credits
    -
    unmatched_debits
)

# =====================================================
# MATCH TYPE SUMMARY
# =====================================================
match_summary = (
    df[
        df["Status"] == "Reconciled"
    ]
    ["Match Type"]
    .value_counts()
    .reset_index()
)

match_summary.columns = [
    "Match Type",
    "Count"
]

# =====================================================
# CREATE SUMMARY TABLE
# =====================================================
summary = pd.DataFrame(
    {
        "Description": [

            "Input File",

            "Total Transactions",

            "Total Debits",

            "Total Credits",

            "Reconciled Transactions",

            "Unreconciled Transactions",

            "Matched Debits",

            "Matched Credits",

            "Unmatched Debits",

            "Unmatched Credits",

            "Outstanding Balance"

        ],


        "Value": [

            os.path.basename(file_path),

            len(df),

            total_debits,

            total_credits,

            len(reconciled),

            len(unreconciled),

            matched_debits,

            matched_credits,

            unmatched_debits,

            unmatched_credits,

            difference

        ]
    }
)

# =====================================================
# OUTPUT FILE
# =====================================================
folder = os.path.dirname(
    file_path
)

filename = os.path.splitext(
    os.path.basename(file_path)
)[0]

output_file = os.path.join(
    folder,
    filename + "_Reconciled.xlsx"
)
# =====================================================
# EXPORT EXCEL
# =====================================================
with pd.ExcelWriter(
    output_file,
    engine="openpyxl"
) as writer:


    reconciled.to_excel(
        writer,
        sheet_name="Reconciled",
        index=False
    )


    unreconciled.to_excel(
        writer,
        sheet_name="Unreconciled",
        index=False
    )


    df.to_excel(
        writer,
        sheet_name="All Transactions",
        index=False
    )


    summary.to_excel(
        writer,
        sheet_name="Summary",
        index=False
    )


    match_summary.to_excel(
        writer,
        sheet_name="Match Analysis",
        index=False
    )

# =====================================================
# FINAL REPORT
# =====================================================
print("=" * 70)

print("RECONCILIATION SUMMARY")

print("=" * 70)

print(
    f"Input File              : "
    f"{os.path.basename(file_path)}"
)

print(
    f"Total Transactions      : "
    f"{len(df):,}"
)

print(
    f"Total Debits            : "
    f"{total_debits:,.2f}"
)

print(
    f"Total Credits           : "
    f"{total_credits:,.2f}"
)

print("-" * 70)

print(
    f"Reconciled Transactions  : "
    f"{len(reconciled):,}"
)

print(
    f"Unreconciled Transactions:"
    f" {len(unreconciled):,}"
)

print("-" * 70)

print(
    f"Matched Debits          : "
    f"{matched_debits:,.2f}"
)

print(
    f"Matched Credits         : "
    f"{matched_credits:,.2f}"
)

print("-" * 70)

print(
    f"Outstanding Debit       : "
    f"{unmatched_debits:,.2f}"
)

print(
    f"Outstanding Credit      : "
    f"{unmatched_credits:,.2f}"
)


print(
    f"Outstanding Balance     : "
    f"{difference:,.2f}"
)

print("-" * 70)

print(
    "MATCH TYPES"
)


print(
    match_summary.to_string(
        index=False
    )
)

print("-" * 70)

print(
    f"Output File:"
)

print(
    output_file
)

print("=" * 70)

print(
    "RECONCILIATION COMPLETE"
)

print("=" * 70)
