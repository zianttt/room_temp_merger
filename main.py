import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill

def auto_detect_sheets_pandas(xls):
    """
    Auto-detect sheets by scanning sheet names for keywords.
    Expected roles: 'Max', 'Min', 'Midband', and 'Room Data'
    """
    role_keywords = {
        "Max": ["max", "maximum"],
        "Min": ["min", "minimum"],
        "Midband": ["midband"],
        "Room Data": ["sensed value", "room"]
    }
    detected = {role: None for role in role_keywords}
    for sheet in xls.sheet_names:
        sheet_lower = sheet.lower()
        for role, keywords in role_keywords.items():
            if any(keyword in sheet_lower for keyword in keywords):
                if detected[role] is None:
                    detected[role] = sheet
    return detected

def process_excel(file) -> BytesIO:
    file.seek(0)
    
    xls = pd.ExcelFile(file)
    detected = auto_detect_sheets_pandas(xls)
    
    room_sheet = detected.get("Room Data")
    min_sheet  = detected.get("Min")
    max_sheet  = detected.get("Max")
    mid_sheet  = detected.get("Midband")  # Optional
    
    missing = []
    for role, sheet in [("Room Data", room_sheet), ("Min", min_sheet), ("Max", max_sheet)]:
        if sheet is None:
            missing.append(role)
    if missing:
        st.error("Missing required sheets: " + ", ".join(missing))
        return None
    if mid_sheet is None:
        st.warning("Sheet for 'Midband' not found. It will be ignored.")
    
    df_room = pd.read_excel(xls, sheet_name=room_sheet, header=None)
    df_min  = pd.read_excel(xls, sheet_name=min_sheet,  header=None)
    df_max  = pd.read_excel(xls, sheet_name=max_sheet,  header=None)
    
    df_room = df_room.dropna(axis=0, how="all").reset_index(drop=True)
    df_room = df_room.dropna(axis=1, how="all").reset_index(drop=True)
    
    df_min = df_min.dropna(axis=0, how="all").reset_index(drop=True)
    df_min = df_min.dropna(axis=1, how="all").reset_index(drop=True)
    
    df_max = df_max.dropna(axis=0, how="all").reset_index(drop=True)
    df_max = df_max.dropna(axis=1, how="all").reset_index(drop=True)
    
    # Define header counts: the first few rows/columns are header (unchanged)
    header_rows_count = 3  # first 3 rows are headers
    header_cols_count = 2  # first 2 columns are headers
    
    # Make a copy for results
    df_result = df_room.copy()
    
    # Process the data region (i.e. rows and columns after the header area)
    # Note: DataFrame indices are 0-based. So data starts at row index header_rows_count and col index header_cols_count.
    for i in range(header_rows_count, df_room.shape[0]):
        for j in range(header_cols_count, df_room.shape[1]):
            # Make sure corresponding cell exists in Min and Max sheets.
            if i < df_min.shape[0] and j < df_min.shape[1] and i < df_max.shape[0] and j < df_max.shape[1]:
                room_val = df_room.iat[i, j]
                min_val  = df_min.iat[i, j]
                max_val  = df_max.iat[i, j]
                try:
                    room_num = float(room_val)
                    min_num  = float(min_val)
                    max_num  = float(max_val)
                except (TypeError, ValueError):
                    # If not numeric, leave the original value.
                    continue
                
                if room_num <= min_num:
                    diff = round(room_num - min_num, 2)
                    df_result.iat[i, j] = f"low: {diff}"
                elif room_num >= max_num:
                    diff = round(room_num - max_num, 2)
                    df_result.iat[i, j] = f"high: {diff}"
                else:
                    df_result.iat[i, j] = "ok"
    
    file.seek(0)
    wb = openpyxl.load_workbook(file)
    
    if "Result" in wb.sheetnames:
        ws_old = wb["Result"]
        wb.remove(ws_old)
    
    ws_result = wb.create_sheet("Result")
    for r_idx, row in enumerate(dataframe_to_rows(df_result, index=False, header=False), start=1):
        for c_idx, value in enumerate(row, start=1):
            ws_result.cell(row=r_idx, column=c_idx, value=value)
    
    # formatting
    blue_fill  = PatternFill(fill_type="solid", start_color="ADD8E6", end_color="ADD8E6")
    green_fill = PatternFill(fill_type="solid", start_color="90EE90", end_color="90EE90")
    red_fill   = PatternFill(fill_type="solid", start_color="FFC7CE", end_color="FFC7CE")
    
    # Apply formatting to the data region only (cells beyond the header area).
    # Since we wrote the DataFrame starting at A1, headers occupy the first header_rows_count rows and header_cols_count columns.
    for row in ws_result.iter_rows(min_row=header_rows_count+1, 
                                   max_row=ws_result.max_row, 
                                   min_col=header_cols_count+1, 
                                   max_col=ws_result.max_column):
        for cell in row:
            if isinstance(cell.value, str):
                if cell.value.startswith("low:"):
                    cell.fill = blue_fill
                elif cell.value.startswith("high:"):
                    cell.fill = red_fill
                elif cell.value == "ok":
                    cell.fill = green_fill
    
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def main():
    st.title("Excel Temperature Checker (Pandas + OpenPyXL)")
    
    uploaded_file = st.file_uploader("Upload your Excel (.xlsx) file", type=["xlsx"])
    if uploaded_file:
        processed_file = process_excel(uploaded_file)
        if processed_file is not None:
            st.success("File processed successfully!")
            st.download_button(
                label="Download updated Excel file",
                data=processed_file,
                file_name="updated_file.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

if __name__ == "__main__":
    main()
