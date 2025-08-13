import openpyxl
from datetime import datetime
import sys


raw_file = "Daily Perf - 12-08-205.xlsx"
template_file = "Daily Performance Sheet - 05 Aug 2025.xlsx"
output_file = "Daily Performance - Filled.xlsx"
SHEETS_TO_IGNORE = ['Home', 'Sheet1']


RAW_HEADER_TO_STANDARD_NAME = {
    "Scheme Name": "Scheme Name",
    "Month End": "Month End",
    "Average Maturity Years": "Avg Maturity",
    "Modified Duration Years": "Mod Duration",
    "YTM (%)": "YTM",
    "Direct Expense Ratio": "Expense Ratio",
    "Latest Date": "Latest Date",
    "Latest NAV(Rs)": "NAV",
    "1 Day": "1 Day",
    "3 Day": "3 Day",
    "1 Week": "1 Week",
    "2 Week": "2 Week",
    "1 Month": "1 Month",
    "3 Months": "3 Months",
    "6 Months": "6 Months",
    "9 Months": "9 Months",
    "1 Year": "1 Year",
    "3 Years": "3 Years",
    "5 Years": "5 Years",
    "10 Years": "10 Years",
    "SINCE INCEPTION": "Since Inception",
    "Cash & Equi": "Cash & Equi",
    "Others": "Others",
    "SOV": "SOV",
    "AA": "AA",
    "AA-": "AA-",
    "AA+": "AA+",
    "AAA/A1+": "AAA/A1+",
    "D": "D",
    "Unrated": "Unrated",
    "Exit Load": "Exit Load",
    "Remark": "Remark",
    "Inception Date": "Inception Date",
    "[Fund Manager 1]": "Fund Manager 1",

}
STANDARD_NAME_TO_TEMPLATE_HEADER = {
    "Scheme Name": "Scheme Name",
    "Month End": "Month End",
    "Avg Maturity": "Average Maturity Years",
    "Mod Duration": "Modified Duration Years",
    "YTM": "YTM (%)",
    "Expense Ratio": "Direct Expense Ratio",
    "Latest Date": "Latest Date",
    "NAV": "Latest NAV(Rs)",
    "1 Day": "1 Day",
    "3 Day": "3 Day",
    "1 Week": "1 Week",
    "2 Week": "2 Week",
    "1 Month": "1 Month",
    "3 Months": "3 Months",
    "6 Months": "6 Months",
    "9 Months": "9 Months",
    "1 Year": "1 Year",
    "3 Years": "3 Years",
    "5 Years": "5 Years",
    "10 Years": "10 Years",
    "Since Inception": "SINCE INCEPTION",
    "Cash & Equi": "Cash & Equi",
    "Others": "Others",
    "SOV": "SOV",
    "AA": "AA",
    "AA-": "AA-",
    "AA+": "AA+",
    "AAA/A1+": "AAA/A1+",
    "D": "D",
    "Unrated": "Unrated",
    "Exit Load": "Exit Load",
    "Remark": "Remark",
    "Inception Date": "Inception Date",
    "Fund Manager 1": "[Fund Manager 1]",
}

print("Starting Data Transfer")
try:
    raw_wb = openpyxl.load_workbook(raw_file, data_only=True)
    template_wb = openpyxl.load_workbook(template_file)
except FileNotFoundError as e:
    print(f"ERROR: Could not find a required file: {e.filename}");
    sys.exit()

latest_date_in_raw_file = None
print("Scanning RAW file to find the latest available AUM date")
for sheet_name in raw_wb.sheetnames:
    if sheet_name in SHEETS_TO_IGNORE: continue
    sheet = raw_wb[sheet_name]
    data_header_row = -1
    for r in range(1, 20):
        for c in sheet[r]:
            if str(c.value).strip() == "Scheme Name": data_header_row = r; break
        if data_header_row != -1: break
    if data_header_row == -1: continue
    available_dates = [c.value for c in sheet[data_header_row] if isinstance(c.value, datetime)]
    if available_dates: latest_date_in_raw_file = max(available_dates); break

if not latest_date_in_raw_file:
    print("ERROR: Could not find any date columns in the headers of the raw file.");
    sys.exit()
master_date = latest_date_in_raw_file
print(f"Dynamically determined latest date from RAW file: {master_date.strftime('%d-%b-%Y')}")


print("Processing and writing data one sheet at a time")
total_rows_written = 0
for sheet_name in raw_wb.sheetnames:
    if sheet_name in SHEETS_TO_IGNORE: continue
    print(f"Processing sheet: '{sheet_name}'")
    raw_sheet = raw_wb[sheet_name]
    if sheet_name not in template_wb.sheetnames:
        print(f"WARNING: Sheet '{sheet_name}' not found in template file. Skipping.");
        continue
    template_sheet = template_wb[sheet_name]


    data_header_row_raw = -1
    for r in range(1, 20):
        for c in raw_sheet[r]:
            if str(c.value).strip() == "Scheme Name": data_header_row_raw = r; break
        if data_header_row_raw != -1: break
    if data_header_row_raw == -1: print(f"WARNING: Could not find header in raw sheet. Skipping."); continue
    raw_col_map = {RAW_HEADER_TO_STANDARD_NAME[str(cell.value).strip()]: col for col, cell in
                   enumerate(raw_sheet[data_header_row_raw], 1) if
                   str(cell.value).strip() in RAW_HEADER_TO_STANDARD_NAME}
    for col, cell in enumerate(raw_sheet[data_header_row_raw], 1):
        if isinstance(cell.value, datetime) and cell.value.date() == master_date.date(): raw_col_map["AUM"] = col; break

    # Find headers and map columns in template sheet
    data_header_row_template = -1
    for r in range(1, 20):
        for c in template_sheet[r]:
            if str(c.value).strip() == "Scheme Name": data_header_row_template = r; break
        if data_header_row_template != -1: break
    if data_header_row_template == -1: print(f"WARNING: Could not find header in template sheet. Skipping."); continue
    dest_col_map = {std_name: col for std_name, tpl_header in STANDARD_NAME_TO_TEMPLATE_HEADER.items() for col, cell in
                    enumerate(template_sheet[data_header_row_template], 1) if str(cell.value).strip() == tpl_header}
    for col, cell in enumerate(template_sheet[data_header_row_template], 1):
        if isinstance(cell.value, datetime) and cell.value.date() == master_date.date(): dest_col_map[
            "AUM"] = col; break


    print("    Clearing old data from template sheet...")
    start_row_to_clear = data_header_row_template + 1
    end_row_to_clear = start_row_to_clear - 1
    scheme_name_dest_col = dest_col_map.get("Scheme Name")

    if scheme_name_dest_col:
        # Find the last row of old data by looking for the first blank cell in the Scheme Name column
        for r_idx in range(start_row_to_clear, template_sheet.max_row + 2):
            if not template_sheet.cell(row=r_idx, column=scheme_name_dest_col).value:
                end_row_to_clear = r_idx - 1
                break


    if end_row_to_clear >= start_row_to_clear:
        print(f"    Clearing old data from row {start_row_to_clear} to {end_row_to_clear}.")
        for row_num in range(start_row_to_clear, end_row_to_clear + 1):
            for dest_col in dest_col_map.values():
                template_sheet.cell(row=row_num, column=dest_col).value = None


    start_row_template = data_header_row_template + 1
    rows_on_this_sheet = 0
    for row_num in range(data_header_row_raw + 1, raw_sheet.max_row + 2):
        scheme_name_val = raw_sheet.cell(row=row_num, column=raw_col_map.get("Scheme Name", 1)).value
        if not scheme_name_val: break

        template_row_index = start_row_template + rows_on_this_sheet
        for std_name, dest_col in dest_col_map.items():
            if std_name in raw_col_map:
                raw_col = raw_col_map[std_name]
                value_to_write = raw_sheet.cell(row=row_num, column=raw_col).value
                template_sheet.cell(row=template_row_index, column=dest_col).value = value_to_write

        rows_on_this_sheet += 1
        total_rows_written += 1
    print(f"    Wrote {rows_on_this_sheet} rows of new data.")

print(f"Total rows written across all sheets: {total_rows_written}")
template_wb.save(output_file)
print(f"Success! Data transferred and saved to {output_file}")