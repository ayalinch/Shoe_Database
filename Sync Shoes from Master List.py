import sqlite3
import pandas as pd
import os
import json

# ── Paths (Make sure these match your Mac!) ────────────────────────
EXCEL_PATH = os.path.expanduser("~/Desktop/WinterLab Master list of footwear.xlsx")
DB_PATH = os.path.expanduser("~/Desktop/Database Final/shoes.db") 

def sync_database():
    if not os.path.exists(EXCEL_PATH):
        print(f"❌ Error: Excel file not found at {EXCEL_PATH}")
        return
    if not os.path.exists(DB_PATH):
        print(f"❌ Error: Database not found at {DB_PATH}")
        return
        
    print("Reading Excel file (this might take a few seconds)...")
    df = pd.read_excel(EXCEL_PATH)
    
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    
    # Ensure the lab_data column exists just in case
    try:
        cur.execute("ALTER TABLE shoes ADD COLUMN lab_data TEXT")
    except:
        pass
        
    cur.execute("SELECT idapt_id, id, notes FROM shoes")
    existing_shoes = {row[0]: {'id': row[1], 'notes': row[2]} for row in cur.fetchall() if row[0]}
    
    updated_count = 0
    inserted_count = 0

    for index, row in df.iterrows():
        idapt = str(row.get('idapt#', '')).strip().upper()
        if not idapt or idapt == 'NAN' or idapt == 'NONE':
            continue
        
        # Core mappings
        brand = str(row.get('Brand Name', '')).strip()
        model = str(row.get('Model#', '')).strip()
        size = str(row.get('Size', '')).strip()
        maa_score = row.get('Final MAA score', None)
        
        if pd.isna(brand) or brand == 'nan': brand = ""
        if pd.isna(model) or model == 'nan': model = ""
        if pd.isna(size) or size == 'nan': size = ""
        
        if pd.isna(maa_score): maa_score = None
        else:
            try: maa_score = float(maa_score)
            except ValueError: maa_score = None

        # Build the exact JSON dictionary for the new Lab Data tab
        lab_dict = {}
        for col in df.columns:
            val = row[col]
            if pd.notna(val) and str(val).strip() != "" and col not in ['idapt#', 'Brand Name', 'Model#', 'Final MAA score']:
                lab_dict[col] = str(val)
                
        lab_data_json = json.dumps(lab_dict)
        
        if idapt in existing_shoes:
            shoe_id = existing_shoes[idapt]['id']
            old_notes = existing_shoes[idapt]['notes'] or ""
            
            # This strips out the old, messy list if you ran the previous script
            if "--- EXCEL LAB DATA ---" in old_notes:
                old_notes = old_notes.split("--- EXCEL LAB DATA ---")[0].strip()
            
            cur.execute("""
                UPDATE shoes 
                SET brand=?, model=?, size=?, maa_mean=?, notes=?, lab_data=?
                WHERE id=?
            """, (brand, model, size, maa_score, old_notes, lab_data_json, shoe_id))
            updated_count += 1
            
        else:
            cur.execute("""
                INSERT INTO shoes (idapt_id, brand, model, size, maa_mean, lab_data)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (idapt, brand, model, size, maa_score, lab_data_json))
            inserted_count += 1

    con.commit()
    con.close()
    
    print("\n✅ Sync Complete!")
    print(f"   Updated {updated_count} existing shoes in the app.")
    print(f"   Added {inserted_count} new shoes from the Excel sheet.")

if __name__ == "__main__":
    sync_database()