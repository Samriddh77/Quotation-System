import pandas as pd
import os

def convert_excel_to_json():
    base_path = os.path.join(os.path.dirname(__file__), 'data')
    excel_path = os.path.join(base_path, 'MASTER PRICE LIST.xlsx')
    json_path = os.path.join(base_path, 'catalog.json')
    
    if not os.path.exists(excel_path):
        print("❌ Error: Could not find 'data/MASTER PRICE LIST.xlsx'")
        return

    print("⏳ Reading Excel file... (This may take a moment)")
    try:
        xls = pd.ExcelFile(excel_path)
        catalog = pd.DataFrame(columns=['Category', 'Item Name', 'List Price', 'Standard Discount', 'GST Rate'])
        
        for sheet in xls.sheet_names:
            process_sheet(xls, sheet, catalog)
            
        # Clean and Save
        catalog['List Price'] = pd.to_numeric(catalog['List Price'], errors='coerce')
        catalog.dropna(subset=['List Price'], inplace=True)
        
        # Save as JSON
        catalog.to_json(json_path, orient='records')
        print(f"✅ Success! Converted {len(catalog)} items to 'data/catalog.json'")
        
    except Exception as e:
        print(f"❌ Error: {e}")

def process_sheet(xls, sheet_name, catalog):
    try:
        df_raw = pd.read_excel(xls, sheet_name, header=None, nrows=50)
    except: return

    header_row = -1
    category = None
    
    # Detect Category
    for i, row in df_raw.iterrows():
        row_str = " ".join([str(x) for x in row if pd.notna(x)])
        if "33 KV" in row_str and "LP" in row_str: category = "HT Cables"
        elif "Copper Armod" in row_str and "LP" in row_str: category = "Copper Armoured"
        elif "Alluminium Armod" in row_str and "LP" in row_str: category = "Aluminium Armoured"
        elif "Code" in row_str and "Price" in row_str: category = "Cosmos Glands"
        elif "Suitable Cable OD" in row_str: category = "HMI Glands"
        elif "Item Description" in row_str and "LP" in row_str:
            if "RR" in row_str or "RR" in sheet_name.upper(): category = "RR Wires"
            elif "LAPP" in row_str or "LAPP" in sheet_name.upper(): category = "LAPP Wires"
        
        if category:
            header_row = i
            break
    
    if header_row != -1 and category:
        df = pd.read_excel(xls, sheet_name, skiprows=header_row)
        cols = [str(c).strip() for c in df.columns]
        df.columns = cols
        
        if category == "HMI Glands": process_hmi(df, catalog)
        elif category == "Cosmos Glands": add_to_catalog(df, 'Code', 'Price', category, catalog)
        elif "Wires" in category:
            lp = next((c for c in cols if c in ['LP', 'COIL LP', 'LP PER MTR']), None)
            if lp: add_to_catalog(df, 'Item Description', lp, category, catalog)
        else:
            item = next((c for c in cols if "33 KV" in c or "Armod" in c), None)
            if item: add_to_catalog(df, item, 'LP', category, catalog)

def add_to_catalog(df, name_col, price_col, cat, catalog):
    temp = pd.DataFrame()
    temp['Item Name'] = df[name_col]
    temp['List Price'] = df[price_col]
    disc = next((c for c in df.columns if 'DISC' in c or 'Discount' in c), None)
    temp['Standard Discount'] = df[disc] if disc else 0
    temp['Category'] = cat
    temp['GST Rate'] = 18.0
    # Append to main catalog
    catalog[cat + "_" + name_col] = 0 # Dummy to ensure concat works
    # Using simple list append for speed in script
    for _, row in temp.iterrows():
        catalog.loc[len(catalog)] = row

def process_hmi(df, catalog):
    # (Simplified HMI logic for script)
    cols = list(df.columns)
    try:
        size = next(i for i, c in enumerate(cols) if "Size (MM)" in c and ".1" not in c)
        rate = next(i for i, c in enumerate(cols) if "RATE PER PCS" in c and ".1" not in c)
        temp = df.iloc[:, [size, rate]].copy()
        temp.columns = ['Item Name', 'List Price']
        temp['Item Name'] = "HMI Single " + temp['Item Name'].astype(str)
        temp['Standard Discount'] = df.iloc[:, rate+1]
        temp['Category'] = 'HMI Glands'
        temp['GST Rate'] = 18.0
        for _, row in temp.iterrows(): catalog.loc[len(catalog)] = row
    except: pass

if __name__ == "__main__":
    convert_excel_to_json()