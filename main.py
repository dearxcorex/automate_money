from openpyxl import load_workbook
from datetime import datetime

from babel.dates import format_date

# Set the locale for Thai language

# file_path = "data\money.xlsx"

# sheet_name = "เบิกจ่าย"  # Use None for all sheets or specify a sheet name like "Sheet1"
# workbook = load_workbook(file_path)

# sheet = workbook[sheet_name]


# # date
# now = datetime.now()
# thai_month_year = format_date(now, "MMMM yyyy", locale="th_TH")
# thai_year = str(int(thai_month_year.split()[-1]) + 543)

# thai_month_year = f"     {thai_month_year.split()[0]} {thai_year}"

# # replace date
# sheet["H7"].value = thai_month_year


# # order
# def order_number(num):

#     order_number = f"อนุมัติ ผภภ. 23  ปฏิบัติการแทน เลขาธิการ กสทช ที่ สทช 2203.3/{num}"
#     sheet["D10"].value = order_number

#     return workbook.save(file_path)


def update_excel(file_path, sheet_name, order_num, date_string):

    # Load the workbook and select the sheet
    workbook = load_workbook(file_path)
    sheet = workbook[sheet_name]

    # Update Thai mounth and year
    now = datetime.now()
    thai_month_year = format_date(now, "MMMM yyyy", locale="th_TH")
    thai_year = str(int(thai_month_year.split()[-1]) + 543)
    thai_month_year = f"     {thai_month_year.split()[0]} {thai_year}"
    sheet["H7"].value = thai_month_year

    # Order number
    order_text = (
        f"อนุมัติ ผภภ. 23  ปฏิบัติการแทน เลขาธิการ กสทช ที่ สทช 2203.3/{order_num}"
    )
    sheet["D10"].value = order_text
    # date order
    sheet["C11"].value = "12/08/2567"
