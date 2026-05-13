# app.py
# ============================================================
# CASHFREE vs ACME RECONCILIATION TOOL
# ============================================================
# FEATURES
# ------------------------------------------------------------
# ✅ Upload XLSX / CSV
# ✅ Exact Match
#     Phone + Amount + Date
#
# ✅ Split Match
#     Phone + Date
#     Multiple ACME rows sum = Cashfree Amount
#
# ✅ Dashboard Metrics
# ✅ Export Excel
# ✅ Export CSV
# ✅ Single Output Sheet
# ✅ Unmatched Records
# ✅ Proper Date Parsing Fix
# ============================================================

import streamlit as st
import pandas as pd
from itertools import combinations
from io import BytesIO

# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Cashfree Reconciliation",
    layout="wide"
)

st.title("💳 Cashfree vs Acme Reconciliation Dashboard")

# ============================================================
# HELPERS
# ============================================================

def read_file(uploaded_file):

    if uploaded_file.name.endswith(".csv"):
        return pd.read_csv(uploaded_file)

    return pd.read_excel(uploaded_file)


# ------------------------------------------------------------
# CLEAN MOBILE
# ------------------------------------------------------------

def clean_mobile(x):

    try:

        x = str(x)

        x = x.replace(".0", "")
        x = x.replace(" ", "")
        x = x.replace("-", "")
        x = x.replace("+91", "")

        # keep only digits
        x = ''.join(filter(str.isdigit, x))

        # last 10 digits
        if len(x) > 10:
            x = x[-10:]

        return x

    except:
        return ""


# ------------------------------------------------------------
# CLEAN AMOUNT
# ------------------------------------------------------------

def clean_amount(x):

    try:
        return round(float(x), 2)

    except:
        return 0.0


# ------------------------------------------------------------
# CLEAN DATE
# IMPORTANT FIX:
# dayfirst=True
# ------------------------------------------------------------

def clean_date(x):

    try:

        return pd.to_datetime(
            x,
            dayfirst=True,
            errors="coerce"
        ).date()

    except:
        return None


# ============================================================
# FILE UPLOAD
# ============================================================

cashfree_file = st.file_uploader(
    "📤 Upload Cashfree File",
    type=["xlsx", "csv"]
)

acme_file = st.file_uploader(
    "📤 Upload Acme File",
    type=["xlsx", "csv"]
)

# ============================================================
# MAIN PROCESS
# ============================================================

if cashfree_file and acme_file:

    # ========================================================
    # READ FILES
    # ========================================================

    cashfree = read_file(cashfree_file)
    acme = read_file(acme_file)

    st.success("✅ Files Uploaded Successfully")

    # ========================================================
    # STANDARDIZE DATA
    # ========================================================

    # -------------------------
    # MOBILE
    # -------------------------

    cashfree["Customer Phone"] = cashfree["Customer Phone"].apply(clean_mobile)

    acme["mobile"] = acme["mobile"].apply(clean_mobile)

    # -------------------------
    # AMOUNT
    # -------------------------

    cashfree["Amount"] = cashfree["Amount"].apply(clean_amount)

    acme["payment_amount"] = acme["payment_amount"].apply(clean_amount)

    # -------------------------
    # DATE
    # IMPORTANT:
    # Use MATCH_DATE
    # -------------------------

    cashfree["MATCH_DATE"] = cashfree["Date"].apply(clean_date)

    acme["MATCH_DATE"] = acme["date_payment"].apply(clean_date)

    # -------------------------
    # USED FLAG
    # -------------------------

    acme["USED"] = False

    # ========================================================
    # DEBUG DATE CHECK
    # ========================================================

    with st.expander("🔍 Debug Date Parsing"):

        st.write("Cashfree Dates")
        st.dataframe(
            cashfree[["Date", "MATCH_DATE"]].head(10)
        )

        st.write("Acme Dates")
        st.dataframe(
            acme[["date_payment", "MATCH_DATE"]].head(10)
        )

    # ========================================================
    # RESULTS STORAGE
    # ========================================================

    results = []

    exact_match_count = 0
    split_match_count = 0
    unmatched_count = 0

    # ========================================================
    # PROCESS CASHFREE ROWS
    # ========================================================

    for cf_idx, cf_row in cashfree.iterrows():

        phone = cf_row["Customer Phone"]

        amount = cf_row["Amount"]

        date = cf_row["MATCH_DATE"]

        matched = False

        # ====================================================
        # 1. EXACT MATCH
        # ====================================================

        exact_matches = acme[

            (acme["mobile"] == phone) &

            (acme["payment_amount"] == amount) &

            (acme["MATCH_DATE"] == date) &

            (acme["USED"] == False)

        ]

        # ----------------------------------------------------
        # EXACT MATCH FOUND
        # ----------------------------------------------------

        if not exact_matches.empty:

            acme_idx = exact_matches.index[0]

            acme.loc[acme_idx, "USED"] = True

            merged = {}

            # Cashfree Columns
            for col in cashfree.columns:

                merged[f"CF_{col}"] = cf_row[col]

            # Acme Columns
            for col in acme.columns:

                merged[f"ACME_{col}"] = exact_matches.iloc[0][col]

            merged["MATCH_TYPE"] = "EXACT MATCH"

            results.append(merged)

            exact_match_count += 1

            matched = True

        # ====================================================
        # 2. SPLIT MATCH
        # ====================================================

        if not matched:

            candidates = acme[

                (acme["mobile"] == phone) &

                (acme["MATCH_DATE"] == date) &

                (acme["USED"] == False)

            ]

            candidate_indices = list(candidates.index)

            found_combo = None

            # ------------------------------------------------
            # TRY COMBINATIONS
            # ------------------------------------------------

            # Try combinations up to 5 rows
            for r in range(2, min(6, len(candidate_indices) + 1)):

                for combo in combinations(candidate_indices, r):

                    combo_sum = round(

                        acme.loc[
                            list(combo),
                            "payment_amount"
                        ].sum(),

                        2
                    )

                    # MATCH FOUND
                    if combo_sum == amount:

                        found_combo = combo

                        break

                if found_combo:
                    break

            # ------------------------------------------------
            # SPLIT MATCH FOUND
            # ------------------------------------------------

            if found_combo:

                combo_rows = acme.loc[list(found_combo)]

                # mark used
                acme.loc[list(found_combo), "USED"] = True

                for _, acme_row in combo_rows.iterrows():

                    merged = {}

                    # Cashfree Columns
                    for col in cashfree.columns:

                        merged[f"CF_{col}"] = cf_row[col]

                    # Acme Columns
                    for col in acme.columns:

                        merged[f"ACME_{col}"] = acme_row[col]

                    merged["MATCH_TYPE"] = "SPLIT MATCH"

                    results.append(merged)

                split_match_count += 1

                matched = True

        # ====================================================
        # 3. UNMATCHED
        # ====================================================

        if not matched:

            merged = {}

            # Cashfree Columns
            for col in cashfree.columns:

                merged[f"CF_{col}"] = cf_row[col]

            merged["MATCH_TYPE"] = "UNMATCHED"

            results.append(merged)

            unmatched_count += 1

    # ========================================================
    # FINAL DATAFRAME
    # ========================================================

    final_df = pd.DataFrame(results)

    # ========================================================
    # DASHBOARD
    # ========================================================

    st.subheader("📊 Dashboard")

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "✅ Exact Matches",
        exact_match_count
    )

    col2.metric(
        "🔗 Split Matches",
        split_match_count
    )

    col3.metric(
        "❌ Unmatched",
        unmatched_count
    )

    st.divider()

    # ========================================================
    # RESULT TABLE
    # ========================================================

    st.subheader("📄 Reconciliation Result")

    st.dataframe(
        final_df,
        use_container_width=True,
        height=600
    )

    # ========================================================
    # EXPORT EXCEL
    # ========================================================

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter"
    ) as writer:

        final_df.to_excel(
            writer,
            index=False,
            sheet_name="Reconciliation"
        )

        workbook = writer.book

        worksheet = writer.sheets["Reconciliation"]

        # Auto column width
        for i, col in enumerate(final_df.columns):

            max_len = max(

                final_df[col]
                .astype(str)
                .map(len)
                .max(),

                len(col)

            )

            worksheet.set_column(
                i,
                i,
                min(max_len + 5, 40)
            )

    excel_data = output.getvalue()

    st.download_button(

        label="⬇ Download Excel",

        data=excel_data,

        file_name="reconciliation_output.xlsx",

        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ========================================================
    # EXPORT CSV
    # ========================================================

    csv_data = final_df.to_csv(
        index=False
    ).encode("utf-8")

    st.download_button(

        label="⬇ Download CSV",

        data=csv_data,

        file_name="reconciliation_output.csv",

        mime="text/csv"
    )

# ============================================================
# END
# ============================================================