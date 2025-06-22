import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

def row_contains_target(df_row, targets):
    """Check if any cell contains target strings (case-insensitive)"""
    for val in df_row.values:
        if any(t.lower() in str(val).lower() for t in targets):
            return True
    return False

def process_bad_rows(wb1_sheets, wb2_sheets):
    # Collect rows from both files that contain "Bad value" or "To be correct"
    targets = ["Bad value", "To be correct"]
    output_rows = []
    headers_added = set()

    for wb_sheets, wb_name in [(wb1_sheets, "wb1"), (wb2_sheets, "wb2")]:
        for sheet_name, df in wb_sheets.items():
            for idx, row in df.iterrows():
                if row_contains_target(row, targets):
                    key = f"{wb_name}:{sheet_name}"
                    # Add headers (row 8, zero-index 8) only once per sheet key (like VBA row 9)
                    if key not in headers_added:
                        if len(df) > 8:
                            output_rows.append(df.iloc[8].tolist())  # row 9 in Excel = index 8 in pandas
                        else:
                            output_rows.append(df.columns.tolist())
                        headers_added.add(key)
                    output_rows.append(row.tolist())
    if output_rows:
        return pd.DataFrame(output_rows)
    else:
        return pd.DataFrame()  # Empty if no bad rows found

def process_two_table(wb1_sheets, wb2_sheets):
    # Merge 2-table sheets side by side with 2 column gap
    if "2-table" not in wb1_sheets or "2-table" not in wb2_sheets:
        return None  # Missing sheets

    df1 = wb1_sheets["2-table"]
    df2 = wb2_sheets["2-table"]
    gap = pd.DataFrame("", index=df1.index, columns=range(2))  # two empty columns gap

    df_out = pd.concat([df1.reset_index(drop=True), gap, df2.reset_index(drop=True)], axis=1)
    return df_out

def compare_three_table(wb1_sheets, wb2_sheets, threshold):
    # Similar logic as VBA CompareThreeTable
    if "3-table" not in wb1_sheets or "3-table" not in wb2_sheets:
        return None

    df1 = wb1_sheets["3-table"].copy()
    df2 = wb2_sheets["3-table"].copy()

    # Build dictionary for markets in df2 keyed by first column (market name)
    dict_markets = {str(row[0]).strip(): idx for idx, row in df2.iterrows() if str(row[0]).strip()}

    colIndexes = [3, 4, 5]  # 0-based for cols 4,5,6 in VBA

    rows_out = []
    # Headers
    headers_1 = df1.columns.tolist()
    headers_2 = df2.columns.tolist()
    out_headers = headers_1 + [""] + headers_2 + ["Status", "% Diff D", "% Diff E", "% Diff F"]
    rows_out.append(out_headers)

    for idx1, row1 in df1.iterrows():
        market = str(row1[0]).strip()
        if market == "":
            continue
        out_row = list(row1)
        out_row.append("")
        if market in dict_markets:
            idx2 = dict_markets[market]
            row2 = df2.iloc[idx2]
            out_row += list(row2)
            diff_found = False
            diffs = []
            for i in colIndexes:
                val1 = row1[i]
                val2 = row2[i]
                try:
                    val1f = float(val1)
                    val2f = float(val2)
                    if val1f != 0:
                        pct_diff = (val2f - val1f) / val1f * 100
                        diffs.append(f"{pct_diff:.2f}%")
                        if abs(pct_diff) > threshold:
                            diff_found = True
                    else:
                        diffs.append("N/A")
                        diff_found = True
                except:
                    diffs.append("N/A")
                    diff_found = True
            out_row.append("Not OK" if diff_found else "OK")
            out_row.extend(diffs)
        else:
            # Market not found in previous
            out_row += [""] * len(df2.columns)
            out_row.append("Not OK")
            out_row.extend(["N/A", "N/A", "N/A"])
        rows_out.append(out_row)

    return pd.DataFrame(rows_out)

def compare_four_table(wb1_sheets, wb2_sheets, threshold):
    if "4-table" not in wb1_sheets or "4-table" not in wb2_sheets:
        return None

    df1 = wb1_sheets["4-table"].copy()
    df2 = wb2_sheets["4-table"].copy()

    # Key: period from column 2 (index 1)
    dict1 = {str(row[1]).strip(): idx for idx, row in df1.iterrows() if str(row[1]).strip()}
    dict2 = {str(row[1]).strip(): idx for idx, row in df2.iterrows() if str(row[1]).strip()}

    all_periods = set(dict1.keys()) | set(dict2.keys())
    cols_to_compare = [3, 4, 5]  # columns D, E, F (0-based 3,4,5)

    output_rows = []
    headers = ["Period", "Curr 1", "Prev 1", "%Diff 1",
               "Curr 2", "Prev 2", "%Diff 2",
               "Curr 3", "Prev 3", "%Diff 3"]
    output_rows.append(headers)

    for period in sorted(all_periods):
        row_out = [period]
        for i in range(3):
            valCurr = "N/A"
            valPrev = "N/A"

            if period in dict1:
                valCurr = df1.iloc[dict1[period], cols_to_compare[i]]
            if period in dict2:
                valPrev = df2.iloc[dict2[period], cols_to_compare[i]]

            row_out.append(valCurr)
            row_out.append(valPrev)

            # Calculate % diff
            try:
                valCurr_f = float(valCurr)
                valPrev_f = float(valPrev)
                if valPrev_f != 0:
                    pct_diff = (valCurr_f - valPrev_f) / valPrev_f * 100
                    pct_diff_str = f"{pct_diff:.2f}%"
                else:
                    pct_diff_str = "N/A"
            except:
                pct_diff_str = "N/A"
            row_out.append(pct_diff_str)
        output_rows.append(row_out)

    df_out = pd.DataFrame(output_rows)
    return df_out

def save_to_excel(bad_rows_df, two_table_df, three_table_df, four_table_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Product sheet = bad rows
        if not bad_rows_df.empty:
            bad_rows_df.to_excel(writer, sheet_name="Product", index=False, header=False)
        else:
            pd.DataFrame(["No bad rows found"]).to_excel(writer, sheet_name="Product", index=False, header=False)

        # 2-table sheet as "Fact"
        if two_table_df is not None:
            two_table_df.to_excel(writer, sheet_name="Fact", index=False, header=False)

        # 3-table sheet as "Markets"
        if three_table_df is not None:
            three_table_df.to_excel(writer, sheet_name="Markets", index=False, header=False)

        # 4-table sheet as "Periods"
        if four_table_df is not None:
            four_table_df.to_excel(writer, sheet_name="Periods", index=False, header=False)

    output.seek(0)
    return output

def read_sheets(uploaded_file):
    # Explicitly specify openpyxl engine for .xlsx files
    try:
        xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
    except ImportError as e:
        st.error("Missing dependency: 'openpyxl'. Please install it with:\n\npip install openpyxl")
        raise e

    sheets = {}
    for sheet_name in xls.sheet_names:
        # Read without header to keep all rows and VBA-style indexing
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None, dtype=object, engine="openpyxl")
        sheets[sheet_name] = df
    return sheets


st.title("Excel Data Comparison Tool")

uploaded_file_current = st.file_uploader("Upload CURRENT Excel file", type=["xlsx"])
uploaded_file_previous = st.file_uploader("Upload PREVIOUS Excel file", type=["xlsx"])

threshold = st.number_input("Enter % threshold to highlight differences", value=5.0, step=0.1)

if uploaded_file_current and uploaded_file_previous:
    if st.button("Run Comparison"):
        with st.spinner("Processing..."):
            wb1_sheets = read_sheets(uploaded_file_current)
            wb2_sheets = read_sheets(uploaded_file_previous)

            bad_rows_df = process_bad_rows(wb1_sheets, wb2_sheets)
            two_table_df = process_two_table(wb1_sheets, wb2_sheets)
            three_table_df = compare_three_table(wb1_sheets, wb2_sheets, threshold)
            four_table_df = compare_four_table(wb1_sheets, wb2_sheets, threshold)

            output_excel = save_to_excel(bad_rows_df, two_table_df, three_table_df, four_table_df)

            st.success("Processing complete!")

            st.download_button(
                label="Download Comparison Report",
                data=output_excel,
                file_name="Comparison_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
