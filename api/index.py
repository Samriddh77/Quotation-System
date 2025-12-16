from flask import Flask, request, jsonify
import os

# Initialize Flask app
app = Flask(__name__)

# Global variable to store our system
# We initialize it as None so the app starts INSTANTLY
quotation_system = None

class QuotationSystem:
    def __init__(self):
        import pandas as pd # Import here to save startup time
        self.catalog = pd.DataFrame(columns=['Category', 'Item Name', 'List Price', 'Standard Discount', 'GST Rate'])
        # Path adjustment for Vercel
        self.base_path = os.path.join(os.path.dirname(__file__), '..', 'data')

    def load_data(self):
        import pandas as pd
        file_path = os.path.join(self.base_path, 'MASTER PRICE LIST.xlsx')
        
        if not os.path.exists(file_path):
            print(f"ERROR: File not found at {file_path}")
            return False

        try:
            xls = pd.ExcelFile(file_path)
            for sheet in xls.sheet_names:
                self.process_sheet(xls, sheet)
            
            self.catalog['List Price'] = pd.to_numeric(self.catalog['List Price'], errors='coerce')
            self.catalog.dropna(subset=['List Price'], inplace=True)
            return True
        except Exception as e:
            print(f"ERROR loading Excel: {e}")
            return False

    def process_sheet(self, xls, sheet_name):
        import pandas as pd
        try:
            df_raw = pd.read_excel(xls, sheet_name, header=None, nrows=50)
        except: return

        header_row_idx = -1
        category = None
        
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
                header_row_idx = i
                break
        
        if header_row_idx != -1 and category:
            df = pd.read_excel(xls, sheet_name, skiprows=header_row_idx)
            cols = [str(c).strip() for c in df.columns]
            df.columns = cols
            
            if category == "HMI Glands": self.process_hmi(df)
            elif category == "Cosmos Glands": self.add_cat(df, 'Code', 'Price', category)
            elif "Wires" in category:
                lp = next((c for c in cols if c in ['LP', 'COIL LP', 'LP PER MTR']), None)
                if lp: self.add_cat(df, 'Item Description', lp, category)
            else:
                item = next((c for c in cols if "33 KV" in c or "Armod" in c), None)
                if item: self.add_cat(df, item, 'LP', category)

    def add_cat(self, df, name_col, price_col, cat):
        import pandas as pd
        temp = pd.DataFrame()
        temp['Item Name'] = df[name_col]
        temp['List Price'] = df[price_col]
        disc = next((c for c in df.columns if 'DISC' in c or 'Discount' in c), None)
        temp['Standard Discount'] = df[disc] if disc else 0
        temp['Category'] = cat
        temp['GST Rate'] = 18.0
        self.catalog = pd.concat([self.catalog, temp], ignore_index=True)

    def process_hmi(self, df):
        import pandas as pd
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
            self.catalog = pd.concat([self.catalog, temp], ignore_index=True)
        except: pass

    def get_quote(self, items, disc_override=None):
        results = []
        for item in items:
            name = item.get('name', '')
            qty = float(item.get('qty', 0))
            mask = self.catalog['Item Name'].astype(str).str.contains(name, case=False, na=False)
            matches = self.catalog[mask]
            
            if matches.empty:
                results.append({'Description': f"NOT FOUND: {name}", 'Total': 0})
                continue
            
            prod = matches.iloc[0]
            disc = float(disc_override) if disc_override else prod['Standard Discount']
            lp = prod['List Price']
            net = lp * (1 - disc/100)
            gst = net * 0.18
            total = (net + gst) * qty
            
            results.append({
                'Description': prod['Item Name'],
                'Qty': qty,
                'ListPrice': round(lp, 2),
                'Discount': round(disc, 2),
                'NetRate': round(net, 2),
                'Total': round(total, 2)
            })
        return results

# ROUTE HANDLER
@app.route('/api/generate', methods=['POST'])
def generate():
    global quotation_system
    # LAZY LOADING: Only load data when the user actually asks for a quote
    if quotation_system is None:
        print("Loading Data for the first time...")
        quotation_system = QuotationSystem()
        success = quotation_system.load_data()
        if not success:
            return jsonify({"error": "Failed to load price list file"}), 500
            
    data = request.json
    return jsonify(quotation_system.get_quote(data.get('items', []), data.get('discount')))

# Simple root route to prevent 404s/500s on basic checks
@app.route('/', methods=['GET'])
def home():
    return "API is Running"