import streamlit as st
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill

def process_excel(file, sheet_mapping) -> BytesIO:
    wb = openpyxl.load_workbook(file)
    
    # Retrieve the sheet names from the mapping
    room_sheet_name = sheet_mapping["Room Data"]
    min_sheet_name = sheet_mapping["Min"]
    max_sheet_name = sheet_mapping["Max"]
    mid_sheet_name = sheet_mapping["Midband"]  # Not used in computation but may be present

    # Check for required sheets in the workbook.
    missing_sheets = []
    for name in [room_sheet_name, min_sheet_name, max_sheet_name]:
        if name not in wb.sheetnames:
            missing_sheets.append(name)
    if missing_sheets:
        st.error("The following required sheets are missing: " + ", ".join(missing_sheets))
        return None
    
    # Warn if the midband sheet (optional) is missing.
    if mid_sheet_name not in wb.sheetnames:
        st.warning(f"Sheet for 'Midband' mapping ({mid_sheet_name}) not found. It will be ignored.")
    
    room_data_sheet = wb[room_sheet_name]
    min_sheet = wb[min_sheet_name]
    max_sheet = wb[max_sheet_name]

    # Remove any existing "Result" sheet to avoid duplicates.
    if "Result" in wb.sheetnames:
        wb.remove(wb["Result"])
    result_sheet = wb.create_sheet("Result")

    # Define fill colors for formatting
    blue_fill = PatternFill(fill_type="solid", start_color="ADD8E6", end_color="ADD8E6")   # blue for low
    green_fill = PatternFill(fill_type="solid", start_color="90EE90", end_color="90EE90")   # green for ok
    red_fill   = PatternFill(fill_type="solid", start_color="FFC7CE", end_color="FFC7CE")   # light red for high

    # Assume that the first 3 rows and first 3 columns are headers.
    max_row = room_data_sheet.max_row
    max_col = room_data_sheet.max_column

    # Process each cell, copying header cells and computing data cells.
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            # Read the cell from Room Data.
            room_cell = room_data_sheet.cell(row=r, column=c)
            # For header cells (first 3 rows or first 3 columns), copy the value.
            if r <= 3 or c <= 3:
                new_value = room_cell.value
            else:
                # Get the corresponding values from the Min and Max sheets.
                room_value = room_cell.value
                min_value = min_sheet.cell(row=r, column=c).value
                max_value = max_sheet.cell(row=r, column=c).value

                # Only perform the calculation if all three values are numeric.
                if all(isinstance(val, (int, float)) for val in [room_value, min_value, max_value]):
                    if room_value <= min_value:
                        diff = round(room_value - min_value, 2)
                        new_value = f"low: {diff}"
                    elif room_value >= max_value:
                        diff = round(room_value - max_value, 2)
                        new_value = f"high: {diff}"
                    else:
                        new_value = "ok"
                else:
                    new_value = room_value

            # Write the new value to the corresponding cell in the Result sheet.
            cell = result_sheet.cell(row=r, column=c, value=new_value)
            
            # For data cells (beyond the header area), apply the desired formatting.
            if r > 3 and c > 3 and isinstance(new_value, str):
                if new_value.startswith("low:"):
                    cell.fill = blue_fill
                elif new_value.startswith("high:"):
                    cell.fill = red_fill
                elif new_value == "ok":
                    cell.fill = green_fill

    # Save the modified workbook into an in-memory buffer.
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def main():
    st.title("Excel Temperature Checker with Formatting and Sheet Mapping")
    st.sidebar.header("Sheet Name Mapping")
    sheet_mapping = {
        "Midband": st.sidebar.text_input("Name for Midband sheet", "Midband"),
        "Min": st.sidebar.text_input("Name for Min sheet", "Min"),
        "Max": st.sidebar.text_input("Name for Max sheet", "Max"),
        "Room Data": st.sidebar.text_input("Name for Room Data sheet", "Room Data")
    }

    uploaded_file = st.file_uploader("Upload your Excel (.xlsx) file", type=["xlsx"])
    
    if uploaded_file:
        processed_file = process_excel(uploaded_file, sheet_mapping)
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
