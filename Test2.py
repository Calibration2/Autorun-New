import smartsheet
import time

# ==========================
# CONFIG
# ==========================
ACCESS_TOKEN = "J52VB2wiLYzdMX7HWciJvsSGQm0zv1mDWKwNd"
SHEET_ID_1 = 5402253021106052  # Sheet1
SHEET_ID_2 = 2896749489246084  # Sheet2

smartsheet_client = smartsheet.Smartsheet(ACCESS_TOKEN)
smartsheet_client.errors_as_exceptions(True)


# ==========================
# HELPER: Get column map
# ==========================
def get_col_map(sheet_id):
    sheet = smartsheet_client.Sheets.get_sheet(sheet_id)
    col_map = {col.title.strip(): col.id for col in sheet.columns}
    return col_map, sheet


# ==========================
# STEP 1: Move E→A, F→B and clear C,D in Sheet1
# ==========================
def step1_move(sheet_id):
    print("🟦 Step1: Move E→A, F→B and clear C,D...")
    col_map, sheet = get_col_map(sheet_id)
    updates = []

    for row in sheet.rows:
        new_row = smartsheet.models.Row()
        new_row.id = row.id
        cell_dict = {}

        val_E = next((c.value for c in row.cells if c.column_id == col_map.get("Sent CAL Date")), None)
        val_F = next((c.value for c in row.cells if c.column_id == col_map.get("Calibration Due Date")), None)

        # E → A
        if val_E not in (None, "") and "ช่องเก็บ Data CAL Date" in col_map:
            cell_dict[col_map["ช่องเก็บ Data CAL Date"]] = smartsheet.models.Cell()
            cell_dict[col_map["ช่องเก็บ Data CAL Date"]].column_id = col_map["ช่องเก็บ Data CAL Date"]
            cell_dict[col_map["ช่องเก็บ Data CAL Date"]].value = val_E

        # F → B
        if val_F not in (None, "") and "ช่องเก็บ Data Calibration Due Date" in col_map:
            cell_dict[col_map["ช่องเก็บ Data Calibration Due Date"]] = smartsheet.models.Cell()
            cell_dict[col_map["ช่องเก็บ Data Calibration Due Date"]].column_id = col_map["ช่องเก็บ Data Calibration Due Date"]
            cell_dict[col_map["ช่องเก็บ Data Calibration Due Date"]].value = val_F

        # Clear C,D
        for col in ["Data CAL Date", "Data Calibration Due Date"]:
            if col in col_map:
                cell_dict[col_map[col]] = smartsheet.models.Cell()
                cell_dict[col_map[col]].column_id = col_map[col]
                cell_dict[col_map[col]].value = ""

        if cell_dict:
            new_row.cells = list(cell_dict.values())
            updates.append(new_row)

    if updates:
        smartsheet_client.Sheets.update_rows(sheet_id, updates)
        print("✅ Step1 completed")
    else:
        print("⚠️ Step1: No data to move or clear")

    time.sleep(5)


# ==========================
# STEP 2: Copy rows with data in Column E → Sheet2
# ==========================
def step2_copy(sheet_src_id, sheet_dst_id):
    print("🟦 Step2: Copy rows with data in Column E → Sheet2...")
    col_map, sheet = get_col_map(sheet_src_id)
    col_E_id = col_map.get("Sent CAL Date")
    if not col_E_id:
        print("❌ Column E not found")
        return

    row_ids = [r.id for r in sheet.rows if any(c.column_id == col_E_id and c.value for c in r.cells)]
    if not row_ids:
        print("⚠️ No rows to copy")
        return

    directive = smartsheet.models.CopyOrMoveRowDirective({
        'row_ids': row_ids,
        'to': {'sheet_id': sheet_dst_id}
    })
    smartsheet_client.Sheets.copy_rows(sheet_src_id, directive)
    print(f"✅ Step2 copied {len(row_ids)} rows to Sheet2")
    time.sleep(5)


# ==========================
# STEP 3: Move A→C, B→D and clear A,B in Sheet2
# ==========================
def step3_move(sheet_id):
    print("🟦 Step3: Move A→C, B→D and clear A,B in Sheet2...")
    col_map, sheet = get_col_map(sheet_id)
    updates = []

    for row in sheet.rows:
        new_row = smartsheet.models.Row()
        new_row.id = row.id
        cell_dict = {}

        val_A = next((c.value for c in row.cells if c.column_id == col_map.get("ช่องเก็บ Data CAL Date")), None)
        val_B = next((c.value for c in row.cells if c.column_id == col_map.get("ช่องเก็บ Data Calibration Due Date")), None)

        # A → C
        if val_A not in (None, "") and "Data CAL Date" in col_map:
            cell_dict[col_map["Data CAL Date"]] = smartsheet.models.Cell()
            cell_dict[col_map["Data CAL Date"]].column_id = col_map["Data CAL Date"]
            cell_dict[col_map["Data CAL Date"]].value = val_A

        # B → D
        if val_B not in (None, "") and "Data Calibration Due Date" in col_map:
            cell_dict[col_map["Data Calibration Due Date"]] = smartsheet.models.Cell()
            cell_dict[col_map["Data Calibration Due Date"]].column_id = col_map["Data Calibration Due Date"]
            cell_dict[col_map["Data Calibration Due Date"]].value = val_B

        # Clear A,B
        for col in ["ช่องเก็บ Data CAL Date", "ช่องเก็บ Data Calibration Due Date"]:
            if col in col_map:
                cell_dict[col_map[col]] = smartsheet.models.Cell()
                cell_dict[col_map[col]].column_id = col_map[col]
                cell_dict[col_map[col]].value = ""

        if cell_dict:
            new_row.cells = list(cell_dict.values())
            updates.append(new_row)

    if updates:
        smartsheet_client.Sheets.update_rows(sheet_id, updates)
        print("✅ Step3 completed")
    else:
        print("⚠️ Step3: No data to move or clear")

    time.sleep(5)


# ==========================
# STEP 4: Change Status 'Complete' → 'The Calibration is still valid' in Sheet2 and move all rows back to Sheet1
# ==========================
def step4_update_status_and_move(sheet_src_id, sheet_dst_id):
    print("🟦 Step4: Update Status in Sheet2 and move rows back to Sheet1...")
    
    # ดึง Sheet Source
    sheet_src = smartsheet_client.Sheets.get_sheet(sheet_src_id)
    col_map_src = {col.title.strip(): col.id for col in sheet_src.columns}
    
    if "Status" not in col_map_src:
        print("❌ Column 'Status' not found in Sheet2")
        return
    
    # อัปเดต Status
    updates = []
    for row in sheet_src.rows:
        old_status = next((c.value for c in row.cells if c.column_id == col_map_src["Status"]), None)
        if str(old_status).strip().lower() == "complete":
            new_row = smartsheet.models.Row()
            new_row.id = row.id
            cell = smartsheet.models.Cell()
            cell.column_id = col_map_src["Status"]
            cell.value = "The Calibration is still valid"
            cell.strict = False
            new_row.cells = [cell]
            updates.append(new_row)

    if updates:
        smartsheet_client.Sheets.update_rows(sheet_src_id, updates)
        print(f"✅ Updated {len(updates)} Status values in Sheet2")
    else:
        print("⚠️ No Status to update in Sheet2")

    time.sleep(2)

    # Move rows back to Sheet1
    row_ids = [r.id for r in sheet_src.rows]
    if row_ids:
        directive = smartsheet.models.CopyOrMoveRowDirective({
            'row_ids': row_ids,
            'to': {'sheet_id': sheet_dst_id}
        })
        smartsheet_client.Sheets.move_rows(sheet_src_id, directive)
        print(f"✅ Moved {len(row_ids)} rows back to Sheet1")
    else:
        print("⚠️ No rows to move")
    
    time.sleep(2)


# ==========================
# STEP 5: Clear Columns E,F in Sheet1 and delete all rows in Sheet2
# ==========================
def step5_clear(sheet1_id, sheet2_id):
    print("🟦 Step5: Clear Columns E,F in Sheet1 and delete all rows in Sheet2...")
    col_map, sheet1 = get_col_map(sheet1_id)

    # Clear E,F
    cols_to_clear = [col_map.get("Sent CAL Date"), col_map.get("Calibration Due Date")]
    updates = []

    for row in sheet1.rows:
        new_row = smartsheet.models.Row()
        new_row.id = row.id
        new_row.cells = []
        for col_id in cols_to_clear:
            if col_id:
                cell = smartsheet.models.Cell()
                cell.column_id = col_id
                cell.value = ""
                cell.strict = False
                new_row.cells.append(cell)
        if new_row.cells:
            updates.append(new_row)

    if updates:
        smartsheet_client.Sheets.update_rows(sheet1_id, updates)
        print("✅ Cleared Columns E,F in Sheet1")
    else:
        print("⚠️ No data to clear in Sheet1")

    # Delete all rows in Sheet2
    sheet2 = smartsheet_client.Sheets.get_sheet(sheet2_id)
    row_ids = [r.id for r in sheet2.rows]
    if row_ids:
        smartsheet_client.Sheets.delete_rows(sheet2_id, row_ids)
        print(f"✅ Deleted all {len(row_ids)} rows in Sheet2")
    else:
        print("⚠️ No rows to delete in Sheet2")

    time.sleep(5)


# ==========================
# MAIN WORKFLOW
# ==========================
if __name__ == "__main__":
    print("\n🚀 Start 5-step workflow...")
    step1_move(SHEET_ID_1)
    step2_copy(SHEET_ID_1, SHEET_ID_2)
    step3_move(SHEET_ID_2)
    step4_update_status_and_move(SHEET_ID_2, SHEET_ID_1)
    step5_clear(SHEET_ID_1, SHEET_ID_2)
    print("\n🎉 Workflow completed!")
