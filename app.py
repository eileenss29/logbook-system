from flask import Flask, render_template, request, jsonify
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import os
import json

app = Flask(__name__)

# --- การตั้งค่า Google Sheets API สำหรับ Vercel ---
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# ดึงค่า JSON จาก Environment Variables ที่เราตั้งชื่อว่า GOOGLE_CREDENTIALS
creds_json = os.environ.get('GOOGLE_CREDENTIALS')

if creds_json:
    # ถ้าอยู่บน Vercel ให้ดึงจากข้อมูลที่ตั้งไว้
    creds_dict = json.loads(creds_json)
    CREDS = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
else:
    # ถ้าไอรีนรันในเครื่องตัวเอง (Local) ให้ใช้ไฟล์ credentials.json เหมือนเดิม
    CREDS = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)

CLIENT = gspread.authorize(CREDS)

# ใส่ ID ของ Google Sheet
SPREADSHEET_ID = "1Zj_tYtz7u0tqk7ESIvEiBv4JAD7cj_iRN6-pqHipAiE"
SPECIAL_DEVICE = "ตู้ปราศจากเชื้อ (Biosafety Cabinet) ESCO"

def get_sheet_by_name(name):
    ss = CLIENT.open_by_key(SPREADSHEET_ID)
    try:
        return ss.worksheet(name.strip())
    except gspread.exceptions.WorksheetNotFound:
        return None

# --- Routes ---
@app.route('/')
def index():
    device_name = request.args.get('device', 'ไม่ระบุชื่อเครื่อง')
    return render_template('index.html', deviceName=device_name)

@app.route('/history')
def history():
    device_name = request.args.get('device', 'ไม่ระบุชื่อเครื่อง')
    return render_template('history.html', deviceName=device_name)

@app.route('/menu')
def menu():
    return render_template('menu.html')

@app.route('/get_history_data')
def get_history_data_api():
    device_name = request.args.get('deviceName', '').strip()
    sheet = get_sheet_by_name(device_name)
    if not sheet:
        return jsonify(None)
    
    is_special = (device_name == SPECIAL_DEVICE)
    all_values = sheet.get_all_values()
    start_idx = 7 if is_special else 5
    num_cols = 10 if is_special else 6
    rows = all_values[start_idx:]
    filtered_data = [row[:num_cols] for row in rows]
    
    return jsonify({
        "isSpecial": is_special,
        "values": filtered_data[::-1]
    })

@app.route('/process_form', methods=['POST'])
def process_form():
    try:
        form_data = request.form
        device_name = form_data.get('deviceName', '').strip()
        sheet = get_sheet_by_name(device_name)
        
        if not sheet:
            return f"ข้อผิดพลาด: ไม่พบชื่อชีท '{device_name}'"

        now = datetime.now()
        day = f"{now.day:02d}"
        month = f"{now.month:02d}"
        year = now.year + 543
        date_for_sheet = f"{day}/{month}/{year}"

        if device_name == SPECIAL_DEVICE:
            all_values = sheet.get_all_values()
            last_row_idx = len(all_values)
            last_date_in_col_a = ""
            if last_row_idx >= 8:
                last_date_in_col_a = all_values[-1][0]

            if form_data.get('formType') == "usage":
                target_row = max(8, last_row_idx + 1)
                time_range = f"{form_data.get('startTime')} - {form_data.get('endTime')} น."
                status = f"ไม่ปกติ: {form_data.get('statusDetail')}" if form_data.get('status') == "อื่นๆ" else "ปกติ"
                row_values = [date_for_sheet, status, form_data.get('job'), time_range, form_data.get('name')]
                # เปลี่ยนจาก .update เป็น .update_cells หรือใช้วิธีที่ gspread รองรับ (บน Vercel แนะนำใช้ append_row หรือ update สั้นๆ)
                sheet.update(range_name=f"A{target_row}:E{target_row}", values=[row_values])
                return "บันทึกการใช้งานเรียบร้อย!"
            else:
                target_row = last_row_idx if last_date_in_col_a == date_for_sheet else max(8, last_row_idx + 1)
                check1 = "/" if form_data.get('check1') else ""
                check2 = "/" if form_data.get('check2') else ""
                check3 = "/" if form_data.get('check3') else ""
                row_values = [date_for_sheet, check1, check2, check3, form_data.get('name')]
                sheet.update(range_name=f"F{target_row}:J{target_row}", values=[row_values])
                return "บันทึกการบำรุงรักษาเรียบร้อย!"
        else:
            time_range = f"{form_data.get('startTime')} - {form_data.get('endTime')} น."
            row_values = [date_for_sheet, time_range, form_data.get('job'), form_data.get('name'), f"'{form_data.get('tel')}", form_data.get('note')]
            sheet.append_row(row_values)
            return "บันทึกข้อมูลเรียบร้อยแล้ว!"

    except Exception as e:
        return f"Error: {str(e)}"

# เพิ่มบรรทัดนี้เพื่อให้ Vercel รู้ว่านี่คือตัวแปรหลัก
app = app

if __name__ == '__main__':
    app.run(debug=True, port=5000)
