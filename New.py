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


# ===============================================================
# Helper: Get column map
# ===============================================================
def get_col_map(sheet_id):
    sheet = smartsheet_client.Sheets.get_sheet(sheet_id)
    return {col.title: col.id for col in sheet.columns}, sheet


# ===============================================================
# Step 1: Move values C→A, D→B and clear C,D
# ===============================================================
def step1_move_and_clear(sheet_id):
    col_map, sheet = get_col_map(sheet_id)
    src_tgt_pairs = [
        ("Data CAL Date", "ช่องเก็บ Data CAL Date"),
        ("Data Calibration Due Date", "ช่องเก็บ Data Calibration Due Date")
    ]
    rows_to_update = []

    for r in sheet.rows:
        new_row = smartsheet.models.Row()
        new_row.id = r.id
        
        for src, tgt in src_tgt_pairs:
            src_id = col_map.get(src)
            tgt_id = col_map.get(tgt)
            if not src_id or not tgt_id:
                continue

            val = next((c.value for c in r.cells if c.column_id == src_id), None)
            if val is not None:

                cell_target = smartsheet.models.Cell()
                cell_target.column_id = tgt_id
                cell_target.value = val
                cell_target.strict = False
                new_row.cells.append(cell_target)
                
                cell_clear = smartsheet.models.Cell()
                cell_clear.column_id = src_id
                cell_clear.value = ""
                cell_clear.strict = False
                new_row.cells.append(cell_clear)

        if new_row.cells:
            rows_to_update.append(new_row)

    if rows_to_update:
        smartsheet_client.Sheets.update_rows(sheet_id, rows_to_update)
        print("✅ Step1: ย้ายข้อมูล C→A, D→B และล้างค่า C,D สำเร็จ")
    else:
        print("⚠️ Step1: ไม่มีข้อมูลให้ย้าย")


# ===============================================================
# Step 2: Copy rows with value in Column F → Sheet2
# ===============================================================
def step2_copy_rows(sheet_id_src, sheet_id_dst):
    col_map, sheet = get_col_map(sheet_id_src)
    f_col_id = col_map.get("Calibration Due Date")
    if not f_col_id:
        print("❌ Step2: ไม่พบคอลัมน์ Calibration Due Date")
        return
    row_ids = [r.id for r in sheet.rows if any(
        c.column_id == f_col_id and c.value for c in r.cells)]
    if not row_ids:
        print("⚠️ Step2: ไม่มีข้อมูลใน Column F")
        return
    smartsheet_client.Sheets.copy_rows(
        sheet_id_src,
        smartsheet.models.CopyOrMoveRowDirective({
            'row_ids': row_ids,
            'to': {'sheet_id': sheet_id_dst}
        })
    )
    print(f"✅ Step2: คัดลอก {len(row_ids)} แถวจาก Sheet1 → Sheet2 สำเร็จ")


import time

# ===============================================================
# Step 3: Clear DataCal, CalDue; move SentCal→DataCal, CalDue→DataCalDue
# ===============================================================
def step3_clean_and_move(sheet_id):
    """
    Step 3:
    - ย้าย Sent Cal Date → Data Cal Date
    - ย้าย Calibration Due Date → Data Calibration Due Date
    - ล้างค่าในช่องเก็บ Data Cal Date / ช่องเก็บ Data Calibration Due Date / Sent Cal Date / Calibration Due Date
    """
    col_map_raw, sheet = get_col_map(sheet_id)

    # สร้าง mapping แบบ case-insensitive
    col_map = {title.strip(): cid for title, cid in col_map_raw.items()}
    lmap = {title.strip().lower(): cid for title, cid in col_map_raw.items()}

    print("Step3 - column titles on sheet:", list(col_map.keys()))

    def find_id_exact_or_substrings(names, substrings=None):
        for name in names:
            cid = lmap.get(name.strip().lower())
            if cid:
                return cid
        if substrings:
            for title_lower, cid in lmap.items():
                if all(s.lower() in title_lower for s in substrings):
                    return cid
        return None

    # 🧭 ระบุชื่อคอลัมน์หลัก
    src_sent = find_id_exact_or_substrings(["Sent Cal Date"])
    src_due = find_id_exact_or_substrings(["Calibration Due Date"])
    tgt_data_cal = find_id_exact_or_substrings(["Data Cal Date"])
    tgt_data_due = find_id_exact_or_substrings(["Data Calibration Due Date"])
    clear_cols = [
        find_id_exact_or_substrings(["ช่องเก็บ Data Cal Date"]),
        find_id_exact_or_substrings(["ช่องเก็บ Data Calibration Due Date"]),
        src_sent,
        src_due
    ]

    # ตรวจสอบคอลัมน์
    for name, cid in {
        "Sent Cal Date": src_sent,
        "Calibration Due Date": src_due,
        "Data Cal Date": tgt_data_cal,
        "Data Calibration Due Date": tgt_data_due
    }.items():
        if not cid:
            print(f"⚠️ ไม่พบคอลัมน์ '{name}'")

    rows_to_update = []

    # 🔄 เริ่มวนอัปเดตข้อมูล
    for r in sheet.rows:
        cell_dict = {}
        new_row = smartsheet.models.Row()
        new_row.id = r.id

        # ✅ ย้าย Sent Cal Date → Data Cal Date
        if src_sent and tgt_data_cal:
            val = next((c.value for c in r.cells if c.column_id == src_sent), None)
            if val not in (None, "", " "):
                cell = smartsheet.models.Cell()
                cell.column_id = tgt_data_cal
                cell.value = val
                cell.strict = False
                cell_dict[cell.column_id] = cell

        # ✅ ย้าย Calibration Due Date → Data Calibration Due Date
        if src_due and tgt_data_due:
            val = next((c.value for c in r.cells if c.column_id == src_due), None)
            if val not in (None, "", " "):
                cell = smartsheet.models.Cell()
                cell.column_id = tgt_data_due
                cell.value = val
                cell.strict = False
                cell_dict[cell.column_id] = cell

        # 🧹 ล้างค่าคอลัมน์ที่กำหนด
        for col_id in [cid for cid in clear_cols if cid]:
            cell_clear = smartsheet.models.Cell()
            cell_clear.column_id = col_id
            cell_clear.value = ""
            cell_clear.strict = False
            cell_dict[cell_clear.column_id] = cell_clear

        if cell_dict:
            new_row.cells = list(cell_dict.values())
            rows_to_update.append(new_row)

    # 💾 อัปเดตแถว
    if rows_to_update:
        smartsheet_client.Sheets.update_rows(sheet_id, rows_to_update)
        print("✅ Step3: ย้ายข้อมูลเรียบร้อย และล้างค่าในคอลัมน์ที่ระบุสำเร็จ")
    else:
        print("⚠️ Step3: ไม่มีข้อมูลให้ย้ายหรือล้าง")

    print("⏳ รอ 5 วินาที ก่อนเริ่ม Step ถัดไป...")
    time.sleep(5)


# ===============================================================
# Step 4: เปลี่ยน Status 'Complete' → 'The Calibration is still valid'
# และย้ายเฉพาะแถวที่มีข้อมูลใน "Data Calibration Due Date" กลับไป Sheet1
# ===============================================================
def step4_status_and_move(source_sheet_id, target_sheet_id):
    print("🟦 Step4: เปลี่ยน Status และย้ายข้อมูลกลับไป Sheet1...")

    # ✅ อ่านข้อมูลทั้งสอง Sheet
    src_sheet = smartsheet_client.Sheets.get_sheet(source_sheet_id).to_dict()
    tgt_sheet = smartsheet_client.Sheets.get_sheet(target_sheet_id).to_dict()

    # ✅ สร้าง mapping ชื่อคอลัมน์ → columnId
    src_col_map = {c['title'].strip(): c['id'] for c in src_sheet['columns']}
    tgt_col_map = {c['title'].strip(): c['id'] for c in tgt_sheet['columns']}

    # ✅ ตรวจสอบว่ามีคอลัมน์ที่จำเป็นไหม
    if "Data Calibration Due Date" not in src_col_map:
        print("❌ ไม่พบคอลัมน์ Data Calibration Due Date ในต้นทาง")
        return
    if "Status" not in tgt_col_map:
        print("❌ ไม่พบคอลัมน์ Status ในปลายทาง")
        return

    data_due_col_id = src_col_map["Data Calibration Due Date"]
    status_col_id = tgt_col_map["Status"]

    # ✅ fallback mapping (ในกรณีชื่อคอลัมน์ต่างกัน)
    fallback_map = {
        "Number": ["Number", "หมายเลข", "เลขที่"],
        "Status": ["Status", "สถานะ"]
    }

    rows_to_copy = []

    for r in src_sheet['rows']:
        # ✅ เลือกเฉพาะแถวที่มีข้อมูลใน Data Calibration Due Date
        data_due_val = next(
            (c.get('value') for c in r['cells'] if c.get('columnId') == data_due_col_id),
            None
        )
        if not data_due_val:
            continue  # ข้ามถ้าไม่มีข้อมูลในคอลัมน์นี้

        new_row = smartsheet.models.Row()
        used_columns = set()

        for c in r['cells']:
            col_title = next((k for k, v in src_col_map.items() if v == c.get('columnId')), None)
            if not col_title:
                continue

            # ✅ mapping จากชื่อคอลัมน์ (ใช้ fallback ด้วย)
            tgt_col_id = tgt_col_map.get(col_title)
            if not tgt_col_id:
                # ลองหา fallback mapping เช่น “Number” → “หมายเลข”
                for main_name, aliases in fallback_map.items():
                    if col_title in aliases and main_name in tgt_col_map:
                        tgt_col_id = tgt_col_map[main_name]
                        break

            if not tgt_col_id or tgt_col_id in used_columns:
                continue

            val = c.get('value')
            if val not in (None, "", " "):
                cell = smartsheet.models.Cell()
                cell.column_id = tgt_col_id
                cell.value = val
                cell.strict = False
                new_row.cells.append(cell)
                used_columns.add(tgt_col_id)

        # ✅ แก้ค่า Status เป็น "The Calibration is still valid"
        found = False
        for cell in new_row.cells:
            if cell.column_id == status_col_id:
                if str(cell.value).strip().lower() == "complete":
                    cell.value = "The Calibration is still valid"
                found = True
                break

        if not found:
            status_cell = smartsheet.models.Cell()
            status_cell.column_id = status_col_id
            status_cell.value = "The Calibration is still valid"
            status_cell.strict = False
            new_row.cells.append(status_cell)

        if new_row.cells:
            rows_to_copy.append(new_row)

    # ✅ เพิ่มกลับไปยัง Sheet1
    if rows_to_copy:
        smartsheet_client.Sheets.add_rows(target_sheet_id, rows_to_copy)
        print(f"✅ Step4: ย้าย {len(rows_to_copy)} แถวกลับไป Sheet1 สำเร็จ (รวมคอลัมน์ Number และแก้ Status แล้ว)")
    else:
        print("⚠️ Step4: ไม่มีข้อมูลที่มีค่าใน Data Calibration Due Date ให้ย้าย")

    # ✅ รอ 5 วินาที ก่อนเริ่ม Step ถัดไป
    print("⏳ รอ 5 วินาที ก่อนเริ่ม Step ถัดไป...")
    time.sleep(5)


import time


# ===============================================================
# Step 5: ลบข้อมูล Sent Cal Date / Calibration Due Date ใน Sheet1
# และลบข้อมูลทั้งหมดใน Sheet2
# ===============================================================
def step5_clear_data(sheet1_id, sheet2_id):
    print("🟦 Step5: เริ่มล้างข้อมูลใน Sheet1 และ Sheet2...")

    # ✅ --- ล้างข้อมูลในคอลัมน์ Sent Cal Date / Calibration Due Date (Sheet1) ---
    col_map, sheet = get_col_map(sheet1_id)
    cols_to_clear = []

    for col_name in ["Sent CAL Date", "Calibration Due Date"]:
        if col_name in col_map:
            cols_to_clear.append(col_map[col_name])
        else:
            print(f"⚠️ ไม่พบคอลัมน์ '{col_name}' ใน Sheet1")

    if cols_to_clear:
        rows_to_update = []
        for r in sheet.rows:
            new_row = smartsheet.models.Row()
            new_row.id = r.id
            for c_id in cols_to_clear:
                cell = smartsheet.models.Cell()
                cell.column_id = c_id
                cell.value = ""
                cell.strict = False
                new_row.cells.append(cell)
            rows_to_update.append(new_row)

        if rows_to_update:
            smartsheet_client.Sheets.update_rows(sheet1_id, rows_to_update)
            print(f"✅ ล้างข้อมูลในคอลัมน์ {', '.join([c for c in ['Sent CAL Date', 'Calibration Due Date'] if c in col_map])} สำเร็จ")
        else:
            print("⚠️ ไม่มีข้อมูลให้ล้างใน Sheet1")
    else:
        print("⚠️ ไม่มีคอลัมน์เป้าหมายให้ล้างใน Sheet1")

    # ✅ --- ลบข้อมูลทั้งหมดใน Sheet2 ---
    try:
        sheet2 = smartsheet_client.Sheets.get_sheet(sheet2_id).to_dict()
        if not sheet2['rows']:
            print("⚠️ Sheet2 ไม่มีแถวให้ลบ")
        else:
            row_ids = [r['id'] for r in sheet2['rows']]
            smartsheet_client.Sheets.delete_rows(sheet2_id, row_ids)
            print(f"✅ ลบข้อมูลทั้งหมด ({len(row_ids)} แถว) ใน Sheet2 สำเร็จ")
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดขณะลบข้อมูลใน Sheet2: {e}")

    # ✅ --- รอ 5 วินาทีก่อนเริ่ม Step ถัดไป ---
    print("⏳ รอ 5 วินาทีก่อนเริ่ม Step ถัดไป...")
    time.sleep(5)

    print("✅ Step5: เสร็จสิ้นการล้างข้อมูลใน Sheet1 และ Sheet2")


# ===============================================================
# MAIN
# ===============================================================
if __name__ == "__main__":
    print("\n🚀 เริ่มรัน Workflow 5 ขั้นตอน...")

    # Step 1
    step1_move_and_clear(SHEET_ID_1)
    time.sleep(5)

    # Step 2
    step2_copy_rows(SHEET_ID_1, SHEET_ID_2)
    time.sleep(5)

    # Step 3
    step3_clean_and_move(SHEET_ID_2)
    time.sleep(5)

    # Step 4
    step4_status_and_move(SHEET_ID_2, SHEET_ID_1)
    time.sleep(5)

    # Step 5
    step5_clear_data(SHEET_ID_1, SHEET_ID_2)
    time.sleep(5)

    print("\n🎉 งานทั้งหมดเสร็จสมบูรณ์แล้ว!")
