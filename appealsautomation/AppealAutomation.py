import csv
import sys
from copy import copy
import openpyxl
from openpyxl.styles import Font
from datetime import datetime
from openpyxl.utils import get_column_letter, column_index_from_string

TEMPLATE = "AppealsTemplate.xlsx"

def create_appeals_workbook(csv_path, row_num):
    csv_path = convert_data_encoding(csv_path)
    wb = openpyxl.Workbook()
    print("✅ -- Workbook created. --")
    create_template_master_sheet(wb, "AppealsTemplate.xlsx")
    #TODO update to take multiple rows
    header, row, col_letter_to_value = create_data_list(csv_path, row_num)
    create_appeal_sheet(wb, row, col_letter_to_value)

    date_str = datetime.now().strftime("%Y-%m-%d")
    wb.save(f"CAP_Workbook_{date_str}.xlsx")


def convert_data_encoding(csv_path):
    with open(csv_path, encoding='latin1') as fin, open('AppealsData_utf8.csv', 'w', encoding='utf-8', newline='') as fout:
        for line in fin:
            fout.write(line)

    print(f"------------------------ NEW WORKBOOK ------------------------")
    print("✅ -- Appeals Data (ugly) converted to UTF-8 (pretty). --")
    return 'AppealsData_utf8.csv'


def create_template_master_sheet(wb, template_path):
    template_wb = openpyxl.load_workbook(template_path)
    template_ws = template_wb["master"]
    ws = wb.active
    ws.title = "master"

    # Copy all cells, including formatting, from template_ws to ws
    for row_idx, row in enumerate(template_ws.iter_rows(), 1):
        for col_idx, cell in enumerate(row, 1):
            new_cell = ws.cell(row=row_idx, column=col_idx, value=cell.value)
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = copy(cell.number_format)
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)
    # Copy column widths
    for col_letter, col_dim in template_ws.column_dimensions.items():
        ws.column_dimensions[col_letter].width = col_dim.width
    # Copy row heights
    for row_num, row_dim in template_ws.row_dimensions.items():
        ws.row_dimensions[row_num].height = row_dim.height

    print("✅ -- Template master sheet created. --")

# OLD 
def create_data_dictionary(csv_path, row_num):
    with open(csv_path, newline='', encoding='utf-8') as f:
        # Skip the first two rows (headers)
        next(f)
        next(f)
        reader = csv.DictReader(f, skipinitialspace=True)
        rows = list(reader)
    if row_num <= 2 or row_num > len(rows) + 3:
        raise ValueError('Row number out of range')
    data = rows[row_num - 1 - 3]  # -1 for 0-based index, - 3 to account for skipped header rows
    print(f"------------------------ NEW SHEET ------------------------")
    print(f"✅ -- Data processed for row: {row_num} --")
    return data

# NEW
def create_data_list(csv_path, row_num):
    with open(csv_path, newline='', encoding='utf-8') as f:
        reader = csv.reader(f)
        all_rows = list(reader)
    # all_rows[0] and all_rows[1] are headers, data starts at all_rows[2]
    if row_num < 1 or row_num > len(all_rows):
        raise ValueError('Row number out of range')
    header_row = all_rows[0]
    data_row = all_rows[row_num - 1]  # row_num is 1-based

    # Build a mapping from Excel column letter to value
    col_letter_to_value = {}
    for idx, value in enumerate(data_row):
        col_letter = get_column_letter(idx + 1)  # Excel columns are 1-based
        col_letter_to_value[col_letter] = value

    print(f"------------------------ NEW SHEET ------------------------")
    print(f"✅ -- Data processed for row: {row_num} --")
    return header_row, data_row, col_letter_to_value

def create_appeal_sheet(wb, data, row_num):
    org_nickname = data["L"].strip().replace(" ", "_")
    # Excel sheet name limit is 31 characters
    sheet_name = org_nickname[:31]
    ws = wb.create_sheet(title=sheet_name)
    copy_template_to_sheet(wb, ws, TEMPLATE)
    print(f"✅ -- Capsheet Created for: {org_nickname} --")

    # Logic to determine both appeal requests and fill appropriate sections
    appeal1 = data["Select Type of Funding for First Appeals Request"]
    appeal2 = data['Select Type of Funding (please choose "N/A" if you would not like to appeal for a second item)']

    fill_header(ws, data)

    # Map funding type to fill function
    funding_type_to_func = {
        "Organizational Maintenance": fill_organizational_maintenance,
        "Stand Alone Program": fill_program1,
        "Series Program": fill_series_program,
        "Stand Alone Trip - Conference/Team Competition": fill_standalone_conference_team_competition,
        "Stand Alone Trip - Other": fill_other_trip,
        "Series Trip - Conference/Team Competition": fill_standalone_trip_competition,  # adjust if needed
        "Series Trip - Other": fill_other_trip,  # adjust if needed
        "Magazine or Journal": fill_media_publication,
        "Newspaper": fill_media_publication,
    }

    # Helper to call the right function for each appeal
    def handle_appeal(appeal_type, appeal_num, split_data):
        func = funding_type_to_func.get(appeal_type)
        if func:
            func(ws, split_data, appeal_num)
        else:
            print(f"⚠️ No fill function mapped for: {appeal_type}")

    data_keys = list(data.keys())
    header_order = data_keys
    hj_idx = header_order.index('Select Type of Funding (please choose "N/A" if you would not like to appeal for a second item)') - 1 
    hk_idx = header_order.index('Select Type of Funding (please choose "N/A" if you would not like to appeal for a second item)')  

    data_appeal1 = {k: data[k] for k in header_order[:hj_idx+1]}
    data_appeal2 = {k: data[k] for k in header_order[hk_idx:]}

    if appeal1 and appeal1 != "...":
        handle_appeal(appeal1, 1, data_appeal1)
    if appeal2 and appeal2 not in ("...", "N/A"):
        handle_appeal(appeal2, 2, data_appeal2)

    fill_footer(ws, data)
    print(f"✅ -- Data inputted. --")


def copy_template_to_sheet(wb, ws, template_path):
    template_wb = openpyxl.load_workbook(template_path)
    # Use the "capsheet" sheet from the template workbook
    template_ws = template_wb["capsheet"]

    # Copy all cells, including formatting, from template_ws to ws
    for row_idx, row in enumerate(template_ws.iter_rows(), 1):
        for col_idx, cell in enumerate(row, 1):
            new_cell = ws.cell(row=row_idx, column=col_idx, value=cell.value)
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = copy(cell.number_format)
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)
    # Copy column widths
    for col_letter, col_dim in template_ws.column_dimensions.items():
        ws.column_dimensions[col_letter].width = col_dim.width
    # Copy row heights
    for row_num, row_dim in template_ws.row_dimensions.items():
        ws.row_dimensions[row_num].height = row_dim.height

def fill_header(ws, data):
    pass
    # ws["A1"] = "Organization Name: " + str(data.get("Recognized Organization Name on the getInvolved Platform", ""))
    # ws["A1"].font = Font(bold=True, size=16)
    # ws["B1"] = "Submitted by: " + data.get("Contact Person Name", "") + " | Email: " + data.get("Contact Email (must be checked daily) ", "") + " | Phone: " + str(data.get("Contact Phone Number", "")) + " | Position: " + data.get("Position", "")
    # ws["B1"].font = Font(bold=True)
    # ws["K1"] = "SABO: " + str(data.get("SABO Account Number:", ""))
    # ws["K1"].font = Font(bold=True, size=16)


# All fill functions now take appeal_num to distinguish between first and second appeal
def fill_program1(ws, data, appeal_num):
    if appeal_num == 1:
        print("----> First Appeal: Standalone Program")
    else:
        fill_program2(ws, data, appeal_num)
    pass

def fill_program2(ws, data, appeal_num):
    if appeal_num == 2:
        print("----> Second Appeal: Standalone Program")
    else:
        print("ERROR: fill_program2 called for appeal_num other than 2")
    pass

def fill_organizational_maintenance(ws, data, appeal_num):
    if appeal_num == 1:
        print("----> First Appeal: Organizational Maintenance")

        # Room Rental and Equipment + Storage (TODO)
        try:
            ws["B24"] = float(data["AC"].strip().replace("$", ""))
        except (ValueError, TypeError):
            ws["B24"] = data["AC"]
        ws["C24"] = data["AD"]

        # Office Supplies
        # try:
        #     ws["B25"] = float(data["Office Supplies:Ê"].strip().replace("$", ""))
        # except (ValueError, TypeError):
        #     ws["B25"] = data["Office Supplies:Ê"]
        # # Only fill description if B25 is not 0 or empty
        # b25_value = ws["B25"].value
        # if b25_value not in (None, "", 0, "0", 0.0):
        #     ws["C25"] = data["Description for Office Supplies:Ê"]

        # # Advertising
        # try:
        #     ws["B26"] = float(data["Advertising:Ê"].strip().replace("$", ""))
        # except (ValueError, TypeError):
        #     ws["B26"] = data["Advertising:Ê For General Meetings only!"]
        # ws["C26"] = data["Description for Advertising:"]

        # # Food for General Meetings
        # try:
        #     ws["B27"] = float(data["Food for General Interest Meetings:"].strip().replace("$", ""))
        # except (ValueError, TypeError):
        #     ws["B27"] = data["Food for General Interest Meetings:"]
        # ws["C27"] = data["Description for Food for General Interest Meetings:"]

        # # Promotional Giveaways
        # try:
        #     ws["B28"] = float(data["Promotional Giveaways:Ê Promotional giveaways must go towards everyone (i.e. we do not fund gift card prizes,but we fund promotional pens that are distributed to everyone)"].strip().replace("$", ""))
        # except (ValueError, TypeError):
        #     ws["B28"] = data["Promotional Giveaways:Ê Promotional giveaways must go towards everyone (i.e. we do not fund gift card prizes,but we fund promotional pens that are distributed to everyone)"]
        # ws["C28"] = data["Description for Promotional Giveaways:"]

        # # Software
        # try:
        #     ws["B30"] = float(data["Software (for University owned computers)/Website (hosting fees):"].strip().replace("$", ""))
        # except (ValueError, TypeError):
        #     ws["B30"] = data["Software (for University owned computers)/Website (hosting fees):"]
        # ws["C30"] = data["Description for Software (for University owned computers)/Website (hosting fees):Ê"]

        # # Duplications
        # try:
        #     ws["B31"] = float(data["Duplications: Copies of programs to be distributed during an event."].strip().replace("$", ""))
        # except (ValueError, TypeError):
        #     ws["B31"] = data["Duplications: Copies of programs to be distributed during an event."]
        # ws["C31"] = data["Description for Duplications:"]

        # # Phone Charges (Film Processing)
        # try:
        #     ws["B32"] = float(data["Film Processing:"].strip().replace("$", ""))
        # except (ValueError, TypeError):
        #     ws["B32"] = data["Film Processing:"]
        # ws["C32"] = data["Description for Film Processing:"]

        # # Uniforms/Costumes
        # try:
        #     ws["B33"] = float(data["Uniforms/Costumes:Ê For performing groups only!"].strip().replace("$", ""))
        # except (ValueError, TypeError):
        #     ws["B33"] = data["Uniforms/Costumes:Ê For performing groups only!"]
        # ws["C33"] = data["Description for Uniforms/Costumes:"]

        # # Other TODO
        # try:
        #     ws["B35"] = float(data["Other:"].strip().replace("$", ""))
        # except (ValueError, TypeError):
        #     ws["B35"] = data["Other:Ê Please specify in the description box what this is for."]
        # ws["C35"] = data["Description for Other:"]

    else:
        print("----> Second Appeal: Organizational Maintenance")

    

def fill_series_program(ws, data, appeal_num):
    if appeal_num == 1:
        print("----> First Appeal: Series Program")
    else:
        print("----> Second Appeal: Series Program")
    pass

def fill_media_publication(ws, data, appeal_num):
    if appeal_num == 1:
        print("----> First Appeal: Media Publication")
    else:
        print("----> Second Appeal: Media Publication")
    pass

def fill_other_trip(ws, data, appeal_num):
    if appeal_num == 1:
        print("----> First Appeal: Other Trip")
    else:
        print("----> Second Appeal: Other Trip")
    pass

def fill_standalone_trip_competition(ws, data, appeal_num):
    if appeal_num == 1:
        print("----> First Appeal: Standalone Trip/Competition")
    else:
        print("----> Second Appeal: Standalone Trip/Competition")
    pass

def fill_standalone_conference_team_competition(ws, data, appeal_num):
    if appeal_num == 1:
        print("----> First Appeal: Standalone Conference/Team Competition")
    else:
        print("----> Second Appeal: Standalone Conference/Team Competition")
    pass

def fill_footer(ws, data):
    pass

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print('Usage: python AppealAutomation.py <AppealsData.csv> <row_number>') #TODO update to take multiple rows
        sys.exit(1)
    csv_path = sys.argv[1]
    row_num = int(sys.argv[2])
    create_appeals_workbook(csv_path, row_num)
