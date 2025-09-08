import openpyxl
from datetime import datetime
import sys
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter

# CONFIGURATION
raw_file = "Daily Performance RAW.xlsx"
template_file = "Daily Performance Sheet - 05 Aug 2025.xlsx"
output_file = "Daily Performance - Filled.xlsx"
SHEETS_TO_IGNORE = ['Home', 'Sheet1']
LOGO_FILENAME = "corpcare_logo.jpg"

RAW_HEADER_TO_STANDARD_NAME = {
    "Scheme Name": "Scheme Name", "Month End": "Month End", "Average Maturity Years": "Avg Maturity",
    "Modified Duration Years": "Mod Duration", "YTM (%)": "YTM", "Direct Expense Ratio": "Expense Ratio",
    "Latest Date": "Latest Date", "Latest NAV(Rs)": "NAV", "1 Day": "1 Day", "3 Day": "3 Day", "1 Week": "1 Week",
    "2 Week": "2 Week", "1 Month": "1 Month", "3 Months": "3 Months", "6 Months": "6 Months", "9 Months": "9 Months",
    "1 Year": "1 Year", "3 Years": "3 Years", "5 Years": "5 Years", "10 Years": "10 Years",
    "SINCE INCEPTION": "Since Inception",
    "Cash & Equi": "Cash & Equi", "Others": "Others", "SOV": "SOV", "AA": "AA", "AA-": "AA-", "AA+": "AA+",
    "AAA/A1+": "AAA/A1+", "D": "D", "Unrated": "Unrated", "Exit Load": "Exit Load", "Remark": "Remark",
    "Inception Date": "Inception Date", "[Fund Manager 1]": "Fund Manager 1",
}

STANDARD_NAME_TO_TEMPLATE_HEADER = {
    "Scheme Name": "Scheme Name", "Month End": "Month End", "Avg Maturity": "Average Maturity Years",
    "Mod Duration": "Modified Duration Years", "YTM": "YTM (%)", "Expense Ratio": "Direct Expense Ratio",
    "Latest Date": "Latest Date", "NAV": "Latest NAV(Rs)", "1 Day": "1 Day", "3 Day": "3 Day", "1 Week": "1 Week",
    "2 Week": "2 Week", "1 Month": "1 Month", "3 Months": "3 Months", "6 Months": "6 Months", "9 Months": "9 Months",
    "1 Year": "1 Year", "3 Years": "3 Years", "5 Years": "5 Years", "10 Years": "10 Years",
    "Since Inception": "SINCE INCEPTION",
    "Cash & Equi": "Cash & Equi", "Others": "Others", "SOV": "SOV", "AA": "AA", "AA-": "AA-", "AA+": "AA+",
    "AAA/A1+": "AAA/A1+", "D": "D", "Unrated": "Unrated", "Exit Load": "Exit Load", "Remark": "Remark",
    "Inception Date": "Inception Date", "Fund Manager 1": "[Fund Manager 1]",
}

#  Global STYLES
title_font = Font(name='Calibri', size=14, bold=True)
date_font = Font(name='Calibri', size=11, color="003366")
button_font = Font(name='Calibri', size=11, color="000000", underline=None)
back_button_font = Font(name='Calibri', size=11, color="000000", bold=True, underline=None)
light_brown_fill = PatternFill(start_color="DCC7A3", end_color="DCC7A3", fill_type="solid")
no_fill = PatternFill(fill_type=None)
heading_fill = PatternFill(start_color="D1B27B", end_color="D1B27B", fill_type="solid")
button_fill = PatternFill(start_color="DCC783", end_color="DCC783", fill_type="solid")
back_button_fill = PatternFill(start_color="DCC783", end_color="DCC783", fill_type="solid")
center_align = Alignment(horizontal='center', vertical='center')
right_align = Alignment(horizontal='right', vertical='center')
thin_border_side = Side(border_style="thin", color="888888")
button_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

# HELPER FUNCTION
def find_header_row(sheet, keyword="Scheme Name"):
    for r in range(1, 20):
        for cell in sheet[r]:
            if str(cell.value).strip() == keyword: return r
    return -1

def update_as_on_date(sheet):
    for r in range(1, 11):
        for cell in sheet[r]:
            if cell.value and str(cell.value).strip().startswith("As on"):
                cell.value = f"As on {datetime.today().strftime('%Y-%b-%d')}";
                return True
    return False

def find_benchmark_row(sheet, keyword="BenchMark"):
    for r in range(1, sheet.max_row + 1):
        cell = sheet.cell(row=r, column=1)
        if cell.value and str(cell.value).strip() == keyword: return r
    return -1

def create_styled_homepage(workbook):
    print("Creating precise styled homepage...")
    if 'Home' in workbook.sheetnames:
        home_sheet = workbook['Home'];
        home_sheet.delete_rows(1, home_sheet.max_row + 1);
        home_sheet.merged_cells.ranges.clear()
    else: home_sheet = workbook.create_sheet('Home', 0)
    try:
        img = Image(LOGO_FILENAME)
        img.height = 105
        img.width = (img.width / img.height) * img.height
        home_sheet.add_image(img, 'A1')
        print(" Company logo added successfully.")
    except FileNotFoundError:
        print(f" Logo file '{LOGO_FILENAME}' not found.")

    home_sheet.merge_cells('H1:L2');
    title_cell = home_sheet['H1'];
    title_cell.value = "Daily Debt MF Tracker";
    title_cell.font = title_font;
    title_cell.fill = heading_fill;
    title_cell.alignment = center_align
    for row in home_sheet['H1:L2']:
        for cell in row: cell.border = button_border
    date_cell = home_sheet['O2'];
    date_cell.value = datetime.today().strftime('%d-%b-%y')
    date_cell.font = date_font;
    date_cell.alignment = right_align
    home_sheet.merge_cells('E4:O5');
    debt_funds_cell = home_sheet['E4'];
    debt_funds_cell.value = "Debt Funds";
    debt_funds_cell.font = title_font;
    debt_funds_cell.fill = heading_fill;
    debt_funds_cell.alignment = center_align
    for row in home_sheet['E4:O5']:
        for cell in row: cell.border = button_border
    data_sheets = sorted([s for s in workbook.sheetnames if s not in SHEETS_TO_IGNORE]);
    start_row, start_col, max_cols = 7, 5, 4;
    button_height, button_width, row_gap, col_gap = 2, 2, 1, 1
    for i, sheet_name in enumerate(data_sheets):
        row_index = i // max_cols;
        col_index = i % max_cols;
        cell_row = start_row + (row_index * (button_height + row_gap));
        cell_col = start_col + (col_index * (button_width + col_gap))
        home_sheet.merge_cells(start_row=cell_row, end_row=cell_row + button_height - 1, start_column=cell_col,
                               end_column=cell_col + button_width - 1)
        button_cell = home_sheet.cell(row=cell_row, column=cell_col);
        button_cell.value = sheet_name;
        button_cell.alignment = center_align
        for r_offset in range(button_height):
            for c_offset in range(button_width):
                cell_to_style = home_sheet.cell(row=cell_row + r_offset, column=cell_col + c_offset);
                cell_to_style.fill = button_fill;
                cell_to_style.border = button_border
        button_cell.hyperlink = f"#'{sheet_name}'!A1";
        button_cell.font = button_font
    home_sheet.sheet_view.showGridLines = False;
    home_sheet.column_dimensions['A'].width = (img.width / 7) if 'img' in locals() else 5
    for i in range(2, 20): home_sheet.column_dimensions[get_column_letter(i)].width = 12
    print("Styled homepage created successfully.")

# MAIN SCRIPT
print("Starting Data Transfer")
try:
    raw_wb = openpyxl.load_workbook(raw_file, data_only=True)
    template_wb = openpyxl.load_workbook(template_file)
except FileNotFoundError as e:
    print(f"Could not find a required file: {e.filename}");
    sys.exit()

latest_date_in_raw_file = None
print("Scanning RAW file to find the latest available AUM date")
for sheet_name in raw_wb.sheetnames:
    if sheet_name in SHEETS_TO_IGNORE: continue
    sheet = raw_wb[sheet_name];
    data_header_row = find_header_row(sheet)
    if data_header_row == -1: continue
    available_dates = [c.value for c in sheet[data_header_row] if isinstance(c.value, datetime)]
    if available_dates: latest_date_in_raw_file = max(available_dates); break
if not latest_date_in_raw_file:
    print("Could not find any date columns in the headers of the raw file.");
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
    update_as_on_date(template_sheet)
    data_header_row_raw = find_header_row(raw_sheet)
    if data_header_row_raw == -1: continue

    raw_col_map = {RAW_HEADER_TO_STANDARD_NAME.get(str(cell.value).strip(), str(cell.value).strip()): col for col, cell
                   in enumerate(raw_sheet[data_header_row_raw], 1)}
    all_raw_dates = sorted([c.value for c in raw_sheet[data_header_row_raw] if isinstance(c.value, datetime)],
                           reverse=True)
    latest_aum_col = next((c for c, h in enumerate(raw_sheet[data_header_row_raw], 1) if
                           isinstance(h.value, datetime) and h.value.date() == all_raw_dates[0].date()),
                          None) if all_raw_dates else None
    older_aum_col = next((c for c, h in enumerate(raw_sheet[data_header_row_raw], 1) if
                          len(all_raw_dates) > 1 and isinstance(h.value, datetime) and h.value.date() == all_raw_dates[
                              1].date()), None)
    older_date_for_legend = all_raw_dates[1] if len(all_raw_dates) > 1 else None

    data_header_row_template = find_header_row(template_sheet)
    if data_header_row_template == -1: continue
    dest_col_map = {std_name: col for std_name, tpl_header in STANDARD_NAME_TO_TEMPLATE_HEADER.items() for col, cell in
                    enumerate(template_sheet[data_header_row_template], 1) if str(cell.value).strip() == tpl_header}
    aum_dest_col = next((c for c, h in enumerate(template_sheet[data_header_row_template], 1) if
                         isinstance(h.value, datetime) or "AUM" in str(h.value)), None)
    if aum_dest_col: dest_col_map['AUM'] = aum_dest_col

    # Clear main data table
    start_row_to_clear = data_header_row_template + 1;
    end_row_to_clear = -1
    scheme_name_dest_col = dest_col_map.get("Scheme Name")
    if scheme_name_dest_col:
        for r_idx in range(start_row_to_clear, template_sheet.max_row + 2):
            if not template_sheet.cell(row=r_idx,
                                       column=scheme_name_dest_col).value: end_row_to_clear = r_idx - 1; break
    if end_row_to_clear >= start_row_to_clear:
        for row_num in range(start_row_to_clear, end_row_to_clear + 1):
            for dest_col in dest_col_map.values():
                template_sheet.cell(row=row_num, column=dest_col).value = None

    # main data with intelligent AUM merging
    start_row_template = data_header_row_template + 1;
    rows_on_this_sheet = 0;
    any_older_data_used = False
    for row_num in range(data_header_row_raw + 1, raw_sheet.max_row + 2):
        if not raw_sheet.cell(row=row_num, column=raw_col_map.get("Scheme Name", 1)).value: break
        template_row_index = start_row_template + rows_on_this_sheet
        for std_name, dest_col in dest_col_map.items():
            if std_name == "AUM":
                aum_value, is_older_data = None, False
                if latest_aum_col and raw_sheet.cell(row=row_num, column=latest_aum_col).value is not None:
                    aum_value = raw_sheet.cell(row=row_num, column=latest_aum_col).value
                elif older_aum_col and raw_sheet.cell(row=row_num, column=older_aum_col).value is not None:
                    aum_value = raw_sheet.cell(row=row_num, column=older_aum_col).value
                    is_older_data, any_older_data_used = True, True
                dest_cell = template_sheet.cell(row=template_row_index, column=dest_col);
                dest_cell.value = aum_value
                dest_cell.fill = light_brown_fill if is_older_data else no_fill
            elif std_name in raw_col_map:
                value_to_write = raw_sheet.cell(row=row_num, column=raw_col_map[std_name]).value
                template_sheet.cell(row=template_row_index, column=dest_col).value = value_to_write
        rows_on_this_sheet += 1;
        total_rows_written += 1
    print(f"Wrote {rows_on_this_sheet} rows of new data.")

    # Benchmark table
    benchmark_start_raw = find_benchmark_row(raw_sheet)
    benchmark_start_template = find_benchmark_row(template_sheet)
    if benchmark_start_raw != -1 and benchmark_start_template != -1:
        print("Found benchmark data, updating template...")
        for r_offset in range(raw_sheet.max_row - benchmark_start_raw + 1):
            raw_row_to_read = benchmark_start_raw + r_offset
            if all(cell.value is None for cell in raw_sheet[raw_row_to_read]): break
            for c_offset, raw_cell in enumerate(raw_sheet[raw_row_to_read], 1):
                dest_row = benchmark_start_template + r_offset;
                dest_col = c_offset
                template_sheet.cell(row=dest_row, column=dest_col).value = raw_cell.value
        print("Benchmark table updated successfully.")

    # Modify Header and Add Legend
        if aum_dest_col:
            top_header_row = data_header_row_template - 2
            top_header_cell = template_sheet.cell(row=top_header_row, column=aum_dest_col)
            mid_header_cell = template_sheet.cell(row=top_header_row + 1, column=aum_dest_col)
            bottom_header_cell = template_sheet.cell(row=data_header_row_template, column=aum_dest_col)
            template_sheet.merge_cells(start_row=top_header_row, end_row=top_header_row + 1, start_column=aum_dest_col,
                                       end_column=aum_dest_col)
            top_header_cell.value = "Corpus"
            top_header_cell.alignment = center_align
            bottom_header_cell.value = "AUM (Cr.)"
        if any_older_data_used and older_date_for_legend:
            benchmark_row = find_benchmark_row(template_sheet) or template_sheet.max_row
            legend_row = benchmark_row + 5
            legend_cell = template_sheet.cell(row=legend_row, column=1)
            legend_cell.fill = light_brown_fill
            template_sheet.column_dimensions['A'].width = 5
            text_cell = template_sheet.cell(row=legend_row, column=2)
            text_cell.value = f"Indicates AUM as of {older_date_for_legend.strftime('%d-%b-%Y')}"
            text_cell.font = Font(bold=True, color="000000")

    back_button = template_sheet['A2']
    back_button.value = "Home";
    back_button.hyperlink = f"#'Home'!A1"
    back_button.font = back_button_font;
    back_button.fill = back_button_fill;
    back_button.alignment = center_align
    back_button.border = button_border
    template_sheet.column_dimensions['A'].width = 20

create_styled_homepage(template_wb)
print(f"Total rows written across all sheets: {total_rows_written}")
template_wb.save(output_file)
print(f"Success! Data transferred and saved to {output_file}")
