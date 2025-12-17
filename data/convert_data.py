import pandas as pd
import os
import re

# --- CONFIGURATION ---
POSSIBLE_INPUTS = [
    "MASTER PRICE LIST.xlsx", 
    "data/MASTER PRICE LIST.xlsx",
    "../MASTER PRICE LIST.xlsx"
]
OUTPUT_FILE = "data/FIXED_MASTER_LIST.xlsx"

def clean_price(val):
    try:
        if pd.isna(val): return 0.0
        s = str(val).strip()
        s_clean = re.sub(r'[^\d.]', '', s)
        return float(s_clean)
    except:
        return 0.0

def clean_coil(val):
    try:
        if pd.isna(val): return 0
        s = str(val).strip()
        s_clean = re.sub(r'[^\d.]', '', s)
        return int(float(s_clean))
    except:
        return 0

def process_master_file():
    # 1. FIND INPUT FILE
    input_path = None
    for p in POSSIBLE_INPUTS:
        if os.path.exists(p):
            input_path = p
            break
            
    if not input_path:
        print(f"❌ Error: 'MASTER PRICE LIST.xlsx' not found!")
        print(f"   Checked locations: {POSSIBLE_INPUTS}")
        return

    print(f"📂 Found file at: {input_path}")
    print(f"   Reading data...")
    
    try:
        xls = pd.ExcelFile(input_path)
    except Exception as e:
        print(f"❌ Error opening Excel file: {e}")
        return

    all_clean_data = []

    for sheet in xls.sheet_names:
        print(f"  👉 Processing sheet: {sheet}")
        
        try:
            df_raw = pd.read_excel(xls, sheet, header=None, nrows=50)
            
            header_idx = -1
            # FIXED: Added 'XLPE', 'KV', 'CORE' to header detection keywords
            for i, row in df_raw.iterrows():
                row_str = " ".join([str(x).upper() for x in row if pd.notna(x)])
                
                has_price = any(x in row_str for x in ["LP", "PRICE", "RATE"])
                has_desc = any(x in row_str for x in ["ITEM", "DESC", "CODE", "SIZE", "CABLE", "PARTICULARS", "XLPE", "KV", "CORE"])
                
                if has_price and has_desc:
                    header_idx = i
                    break
            
            if header_idx == -1:
                print(f"     ⚠️  Skipping {sheet} (No valid header found)")
                continue

            # Reload with correct header
            df = pd.read_excel(xls, sheet, skiprows=header_idx)
            df.columns = [str(c).strip() for c in df.columns]
            
            # --- INTELLIGENT COLUMN MAPPING ---
            
            # 1. Description
            desc_col = None
            # FIXED: Added 'XLPE', 'KV', 'CORE' to column name candidates
            possible_desc = ["Item Description", "Description", "Particulars", "Code", "Size", "Item Name", "XLPE", "KV", "CORE"]
            
            for col in df.columns:
                if "CABLE" in col.upper() and "ARM" in col.upper(): possible_desc.append(col)
                
            for candidate in possible_desc:
                match = next((c for c in df.columns if candidate.upper() in c.upper()), None)
                if match: desc_col = match; break
            
            if not desc_col: desc_col = df.columns[0] # Fallback to first column

            # 2. Price
            price_col = None
            price_col = next((c for c in df.columns if "PER MTR" in c.upper() or "PER METER" in c.upper()), None)
            if not price_col:
                for keyword in ["LP", "List Price", "Rate", "Price", "Amount", "MRP"]:
                    match = next((c for c in df.columns if keyword.upper() == c.upper() or keyword.upper() in c.upper()), None)
                    if match and "AMOUNT" in match.upper() and any("RATE" in x.upper() for x in df.columns): continue
                    if match: price_col = match; break

            # 3. Discount
            disc_col = next((c for c in df.columns if "DISC" in c.upper() or "OFF" in c.upper()), None)

            # 4. Coil Length
            coil_col = next((c for c in df.columns if "COIL" in c.upper() and ("LEN" in c.upper() or "MTR" in c.upper()) and "QTY" not in c.upper()), None)

            if desc_col and price_col:
                clean_df = pd.DataFrame()
                clean_df["Item Description"] = df[desc_col].astype(str)
                clean_df["List Price"] = df[price_col].apply(clean_price)
                
                if disc_col:
                    clean_df["Standard Discount"] = pd.to_numeric(df[disc_col], errors='coerce').fillna(0)
                else:
                    clean_df["Standard Discount"] = 0
                
                if coil_col:
                    clean_df["Coil Length (Mtr)"] = df[coil_col].apply(clean_coil)
                else:
                    clean_df["Coil Length (Mtr)"] = 0
                
                s_up = sheet.upper()
                p_up = price_col.upper()
                if "MTR" in p_up or "METER" in p_up: uom = "Mtr"
                elif "PC" in p_up or "PIECE" in p_up: uom = "Pc"
                elif "GLAND" in s_up or "HMI" in s_up or "COSMOS" in s_up: uom = "Pc"
                else: uom = "Mtr" 
                
                clean_df["UOM"] = uom

                # Filter bad rows
                clean_df = clean_df[clean_df["List Price"] > 0]
                clean_df = clean_df[clean_df["Item Description"] != "nan"]
                
                all_clean_data.append((sheet, clean_df))
                print(f"     ✅ Mapped {len(clean_df)} items.")
            else:
                print(f"     ⚠️  Skipping {sheet} (Columns missing)")
        
        except Exception as e:
            print(f"     ❌ Error reading sheet: {e}")

    if all_clean_data:
        os.makedirs("data", exist_ok=True)
        with pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter') as writer:
            for sheet_name, df_clean in all_clean_data:
                safe_name = sheet_name[:30].replace("/", "-") 
                df_clean.to_excel(writer, sheet_name=safe_name, index=False)
                
        print(f"\n🎉 Success! Created '{OUTPUT_FILE}'")
    else:
        print("\n❌ Failed to extract data.")

if __name__ == "__main__":
    process_master_file()