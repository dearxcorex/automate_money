from openpyxl import load_workbook
from datetime import datetime

from babel.dates import format_date

# Set the locale for Thai language

file_path = "data/money.xlsx"

#member 
#create dict 

team = {
    "member1": {"name":"นายภูวกฤต  พลชิงชัย","position":"นตป.ก2","allowance":350},
    "member2": {"name":"นางสาวธันยพัฒน์ ภาดาเพิ่มผลสมบัติ","position":"นตป.ก1","allowance":350},
    "member3":{ "name":"นายธนกฤต ชื่นฉิมพลี","position":"นตป.ก1","allowance":300},
}



# Define the Thai to Gregorian year conversion function
def thai_to_gregorian(thai_date):
    day, month_thai, year_thai = thai_date.split()
    thai_months = {
        "มกราคม": 1, "กุมภาพันธ์": 2, "มีนาคม": 3, "เมษายน": 4,
        "พฤษภาคม": 5, "มิถุนายน": 6, "กรกฎาคม": 7, "สิงหาคม": 8,
        "กันยายน": 9, "ตุลาคม": 10, "พฤศจิกายน": 11, "ธันวาคม": 12,
    }
    gregorian_year = int(year_thai) - 543
    month = thai_months[month_thai]
    return f"{gregorian_year}-{month:02d}-{int(day):02d}"




def update_excel(file_path, order_num, date_string,start_day,stop_day,provice,start_time):

    # Load the workbook and select the sheet
    workbook = load_workbook(file_path)
    sheet_1 = workbook['เบิกจ่าย']
    sheet_2 = workbook['หลักฐาน'] 
    print(sheet_2["A1"].value)
    # Update Thai mounth and year
    now = datetime.now()
    thai_month_year = format_date(now, "MMMM yyyy", locale="th_TH")
    thai_year = str(int(thai_month_year.split()[-1]) + 543)
    thai_month_year = f"     {thai_month_year.split()[0]} {thai_year}"
    sheet_1["H7"].value = thai_month_year

    # Order number
    order_text = (
        f"อนุมัติ ผภภ. 23  ปฏิบัติการแทน เลขาธิการ กสทช ที่ สทช 2203.3/{order_num}"
    )
    sheet_1["D10"].value = order_text



    # date order
    day,month,year = map(int,date_string.split('/'))
    parsed_data = datetime(year,month,day)
   #format thai date 
    thai_date = format_date(parsed_data, "d MMMM yyyy", locale="th_TH") 
    sheet_1["C11"].value = thai_date

    #where you work
    work_info = (
        f"เพื่อปฏิบัติงานนอกที่ตั้ง ระหว่างวันที่ {start_day} - {stop_day} {thai_month_year.strip()} ในพื้นที่จังหวัด{provice}  และพื้นที่ใกล้เคียง"
    )

    sheet_1["C15"] = work_info

    #day start and day stop 
    day_start_to_work = f"{start_day} {thai_month_year.strip()}" 
    sheet_1["F19"] = day_start_to_work
    sheet_1["J19"] = f"{str(start_time)} น."

    day_stop_to_work = f"{stop_day} {thai_month_year.strip()}"
    sheet_1["G20"] = day_stop_to_work
    sheet_1["J20"] = f"{str(stop_day)} น."

    start_datetime = datetime.strptime(thai_to_gregorian(day_start_to_work) + " " + start_time, "%Y-%m-%d %H:%M")
    stop_datetime = datetime.strptime(thai_to_gregorian(day_stop_to_work) + " " + stop_time, "%Y-%m-%d %H:%M")

    delta = stop_datetime - start_datetime

    #Extract days, hours and minutes without seconds
    total_minutes = delta.seconds // 60
    days = delta.days 
    hours = int(total_minutes // 60) % 24 
    minutes = int(total_minutes % 60)

    # print(days,hours,minutes)
    sheet_1["D21"] = days
    sheet_1["F21"] = hours
    sheet_1["H21"] = minutes

order_num =123
date_string = "12/08/2567"
start_day = "20"
stop_day = "21"
provice = 'บุรีรัมย์'
start_time = "09:00"
stop_time = "17:45"


update_excel(file_path,order_num,date_string,start_day,stop_day,provice,start_time)