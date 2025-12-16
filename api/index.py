from flask import Flask, request, jsonify
import pandas as pd
import os

app = Flask(__name__)

# Load Data Once (Super Fast with JSON)
catalog = pd.DataFrame()
try:
    json_path = os.path.join(os.path.dirname(__file__), '..', 'data', 'catalog.json')
    if os.path.exists(json_path):
        catalog = pd.read_json(json_path, orient='records')
        print(f"Loaded {len(catalog)} items from JSON.")
    else:
        print("Error: catalog.json not found!")
except Exception as e:
    print(f"Error loading JSON: {e}")

@app.route('/api/generate', methods=['POST'])
def generate():
    data = request.json
    items = data.get('items', [])
    discount_override = data.get('discount', None)
    
    results = []
    if catalog.empty:
        return jsonify([{"Description": "System Error: Data not loaded", "Total": 0}])

    for item in items:
        name = item.get('name', '')
        qty = float(item.get('qty', 0))
        
        # Fast Search
        mask = catalog['Item Name'].astype(str).str.contains(name, case=False, na=False)
        matches = catalog[mask]
        
        if matches.empty:
            results.append({'Description': f"NOT FOUND: {name}", 'Total': 0})
            continue
        
        prod = matches.iloc[0]
        disc = float(discount_override) if discount_override else prod['Standard Discount']
        lp = prod['List Price']
        net = lp * (1 - disc/100)
        gst = net * 0.18
        total = (net + gst) * qty
        
        results.append({
            'Category': prod['Category'],
            'Description': prod['Item Name'],
            'Qty': qty,
            'ListPrice': round(lp, 2),
            'Discount': round(disc, 2),
            'NetRate': round(net, 2),
            'Total': round(total, 2)
        })
        
    return jsonify(results)