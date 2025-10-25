import smartsheet
import time

# ==========================
# CONFIG
# ==========================
ACCESS_TOKEN = "J52VB2wiLYzdMX7HWciJvsSGQm0zv1mDWKwNd"
SHEET_ID_1 = 5402253021106052   # Sheet1 ID
SHEET_ID_2 = 2896749489246084   # Sheet2 ID

smartsheet_client = smartsheet.Smartsheet(ACCESS_TOKEN)
smartsheet_client.errors_as_exceptions(True)


print("🚀 เริ่มรัน Workflow Smartsheet Auto...\n")

# ===============================================================
# Step 1: ตรวจสอบและย้ายข้อมูลจาก Sheet1 → Sheet2
# ===============================================================
def step1_move_to_sheet2(source_sheet_id, target_sheet_id):
    print("🟦 Step1: ตรวจสอบและย้ายข้อมูลจาก Sheet1 → Sheet2")
    src_sheet = smartsheet_client.Sheets.get_sheet(source_sheet_id)
    tgt_sheet = smartsheet_client.Sheets.get_sheet(target_sheet_id)

    if not src_sheet.rows:
        print("⚠️ ไม่มีข้อมูลใน Sheet1")
        return

    rows_to_copy = []
    src_col_map = {c.title.strip(): c.id for c in src_sheet.columns}
    tgt_col_map = {c.title.strip(): c.id for c in tgt_sheet.columns}

    # ตัวอย่าง: ย้ายเฉพาะแถวที่ Status = "Overdue"
    for r in src_sheet.rows:
        status_val = next((c.value for c in r.cells if c.column_id == src_col_map.get("Status")), None)
        if str(status_val).strip().lower() == "overdue":
            new_row = smartsheet.models.Row()
            for c in r.cells:
                col_title = next((t for t, cid in src_col_map.items() if cid == c.column_id), None)
                tgt_col_id = tgt_col_map.get(col_title)
                if tgt_col_id and c.value not in (None, "", " "):
                    new_row.cells.append({
                        "column_id": tgt_col_id,
                        "value": c.value,
                        "strict": False
                    })
            rows_to_copy.append(new_row)

    if rows_to_copy:
        smartsheet_client.Sheets.add_rows(target_sheet_id, rows_to_copy)
        print(f"✅ ย้ายข้อมูล {len(rows_to_copy)} แถวสำเร็จ")
    else:
        print("⚠️ ไม่มีข้อมูลให้ย้าย")

    print("⏳ รอ 5 วินาที...")
    time.sleep(5)

# ===============================================================
# Step 2: ตรวจสอบคอลัมน์ Data Calibration Due Date
# ===============================================================
def step2_check_column(sheet_id):
    print("🟦 Step2: ตรวจสอบ Column Data Calibration Due Date")
    sheet = smartsheet_client.Sheets.get_sheet(sheet_id)
    columns = [c.title for c in sheet.columns]
    if "Data Calibration Due Date" not in columns:
        print("⚠️ ไม่พบคอลัมน์ Data Calibration Due Date")
    else:
        print("✅ พบคอลัมน์ Data Calibration Due Date")
    time.sleep(2)

# ===============================================================
# Step 3: ตรวจสอบข้อมูลในคอลัมน์ที่ต้องการ
# ===============================================================
def step3_verify_data(sheet_id):
    print("🟦 Step3: ตรวจสอบข้อมูลใน Sheet2")
    sheet = smartsheet_client.Sheets.get_sheet(sheet_id)
    if not sheet.rows:
        print("⚠️ ไม่มีข้อมูลใน Sheet2")
        return

    col_map = {c.title.strip(): c.id for c in sheet.columns}
    if "Calibration Due Date" not in col_map:
        print("⚠️ ไม่มีคอลัมน์ Calibration Due Date")
        return

    has_data = any(
        c.value for r in sheet.rows for c in r.cells if c.column_id == col_map["Calibration Due Date"]
    )
    if has_data:
        print("✅ พบข้อมูล Calibration Due Date")
    else:
        print("⚠️ ไม่มีข้อมูล Calibration Due Date")

    time.sleep(3)

# ===============================================================
# Step 4: เปลี่ยน Status + ย้ายกลับไป Sheet1
# ===============================================================
def step4_status_and_move(source_sheet_id, target_sheet_id):
    print("🟦 Step4: เปลี่ยน Status และย้ายข้อมูลกลับไป Sheet1...")

    src_sheet = smartsheet_client.Sheets.get_sheet(source_sheet_id)
    tgt_sheet = smartsheet_client.Sheets.get_sheet(target_sheet_id)

    if not src_sheet.rows:
        print("⚠️ ไม่มีข้อมูล rows ใน Sheet ต้นทาง (อาจเป็น Sheet ว่าง หรือ ID ผิด)")
        return

    src_col_map = {c.title.strip(): c.id for c in src_sheet.columns}
    tgt_col_map = {c.title.strip(): c.id for c in tgt_sheet.columns}

    if "Data Calibration Due Date" not in src_col_map:
        print("❌ ไม่พบคอลัมน์ Data Calibration Due Date")
        return
    if "Status" not in tgt_col_map:
        print("❌ ไม่พบคอลัมน์ Status")
        return

    data_due_col_id = src_col_map["Data Calibration Due Date"]
    status_col_id = tgt_col_map["Status"]
    rows_to_copy = []

    for r in src_sheet.rows:
        data_due_val = next((c.value for c in r.cells if c.column_id == data_due_col_id), None)
        if not data_due_val:
            continue

        new_row = smartsheet.models.Row()
        for c in r.cells:
            col_title = next((t for t, cid in src_col_map.items() if cid == c.column_id), None)
            tgt_col_id = tgt_col_map.get(col_title)
            if tgt_col_id and c.value not in (None, "", " "):
                new_row.cells.append({
                    "column_id": tgt_col_id,
                    "value": c.value,
                    "strict": False
                })

        new_row.cells.append({
            "column_id": status_col_id,
            "value": "The Calibration is still valid",
            "strict": False
        })
        rows_to_copy.append(new_row)

    if rows_to_copy:
        smartsheet_client.Sheets.add_rows(target_sheet_id, rows_to_copy)
        print(f"✅ Step4: ย้ายกลับ {len(rows_to_copy)} แถวสำเร็จ")
    else:
        print("⚠️ ไม่มีข้อมูลที่มีค่าใน Data Calibration Due Date ให้ย้าย")

    print("⏳ รอ 5 วินาที...")
    time.sleep(5)

# ===============================================================
# Step 5: สรุปผล
# ===============================================================
def step5_summary():
    print("🟩 Step5: กระบวนการทั้งหมดเสร็จสมบูรณ์ ✅")
    print("===================================================")

# ===============================================================
# 🔁 รันทุก Step ตามลำดับ
# ===============================================================
def main():
    print("🚀 เริ่ม Workflow 5 ขั้นตอน...\n")
    step1_move_to_sheet2(SHEET_ID_1, SHEET_ID_2)
    step2_check_column(SHEET_ID_2)
    step3_verify_data(SHEET_ID_2)
    step4_status_and_move(SHEET_ID_2, SHEET_ID_1)
    step5_summary()

if __name__ == "__main__":
    main()
    
