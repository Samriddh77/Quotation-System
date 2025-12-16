import streamlit as st
import pandas as pd
import os
import re

# --- PAGE CONFIG ---
st.set_page_config(page_title="Quotation Generator", layout="wide")

# --- 1. HELPERS ---
def clean_price_value(val):
    """ Cleans price string '5,355' -> 5355.0 """
    if pd.isna(val): return None
    s = str(val).strip()
    s_clean = re.sub(r'[^\d.]', '', s)
    try:
        return float(s_clean)
    except:
        return None

def detect_unit(sheet_name, col_name):
    """
    Decides if the Unit is 'Mtr', 'Pc', or 'Coil' based on 
    the Sheet Name and the Column Name used for price.
    """
    s_up = sheet_name.upper()
    c_up = col_name.upper()
    
    # 1. Explicit Column Name Wins
    if "MTR" in c_up or "METER" in c_up:
        return "Mtr"
    if "COIL" in c_up:
        return "Coil"
    if "PC" in c_up or "PIECE" in c_up:
        return "Pc"
        
    # 2. Category Heuristics
    if "GLAND" in s_up or "COSMOS" in s_up or "HMI" in s_up:
        return "Pc"
    if "CABLE" in s_up or "WIRE" in s_up or "ARM" in s_up:
        # Cables are almost always per Meter unless specified otherwise
        return "Mtr"
        
    return "Unit" # Fallback

# --- 2. DATA LOADER ---
@st.cache_data(show_spinner=True)
def load_data_robust():
    possible_paths = [
        "data/MASTER PRICE LIST.xlsx",
        "MASTER PRICE LIST.xlsx",
        "data/master price list.xlsx"
    ]
    file_path = next((p for p in possible_paths if os.path.exists(p)), None)
    
    if not file_path:
        return None, ["❌ 'MASTER PRICE LIST.xlsx' not found!"]

    xls = pd.ExcelFile(file_path)
    all_data = []
    logs = []

    for sheet_name in xls.sheet_names:
        try:
            # Read chunk
            df_raw = pd.read_excel(xls, sheet_name, header=None, nrows=60)
            
            # A. FIND HEADER
            header_idx = -1
            for i, row in df_raw.iterrows():
                row_str = " ".join([str(x).upper() for x in row if pd.notna(x)])
                if (("LP" in row_str or "PRICE" in row_str or "RATE" in row_str) and 
                    ("ITEM" in row_str or "DESC" in row_str or "CODE" in row_str or "SIZE" in row_str)):
                    header_idx = i
                    break
            
            if header_idx == -1:
                # Fallback
                for i, row in df_raw.iterrows():
                    if row.count() > 3:
                        header_idx = i
                        break

            # B. READ SHEET
            df = pd.read_excel(xls, sheet_name, skiprows=header_idx)
            df.columns = [str(c).strip() for c in df.columns]
            
            # C. IDENTIFY COLUMNS
            name_col = None
            price_col = None
            disc_col = None
            
            # --- PRIORITY LOGIC FOR PRICE ---
            # 1. "Per Meter" Preference for Wires/Cables
            for col in df.columns:
                c_up = col.upper()
                if "MTR" in c_up or "METER" in c_up:
                    if any(x in c_up for x in ['LP', 'RATE', 'PRICE']):
                        price_col = col
                        break
            
            # 2. Standard Price Column
            if not price_col:
                for col in df.columns:
                    c_up = col.upper()
                    if any(x in c_up for x in ['LP', 'PRICE', 'RATE', 'MRP', 'LIST', 'AMOUNT']):
                        if "AMOUNT" in c_up and any("RATE" in c.upper() for c in df.columns): continue
                        price_col = col
                        break

            # Find Name & Discount
            for col in df.columns:
                c_up = col.upper()
                if not name_col and any(x in c_up for x in ['DESC', 'ITEM', 'PARTICULARS', 'CODE', 'SIZE', 'TYPE', 'CABLE']):
                    name_col = col
                if not disc_col and any(x in c_up for x in ['DISC', 'OFF']):
                    disc_col = col
            
            if not name_col: name_col = df.columns[0]
            
            # D. PROCESS DATA
            if name_col and price_col:
                temp = pd.DataFrame()
                temp['Item Name'] = df[name_col].astype(str)
                temp['List Price'] = df[price_col].apply(clean_price_value)
                
                if disc_col:
                    temp['Standard Discount'] = pd.to_numeric(df[disc_col], errors='coerce').fillna(0)
                else:
                    temp['Standard Discount'] = 0
                
                temp['Category'] = sheet_name
                
                # --- DETECT UNIT ---
                detected_unit = detect_unit(sheet_name, price_col)
                temp['Unit'] = detected_unit
                
                # Filter valid
                temp.dropna(subset=['List Price'], inplace=True)
                temp = temp[temp['List Price'] > 0]
                
                # HMI Special Handling
                if "HMI" in sheet_name.upper() and len(df.columns) > 8:
                    try:
                        r_name_c = [c for c in df.columns[6:] if "SIZE" in str(c).upper()]
                        r_price_c = [c for c in df.columns[6:] if "RATE" in str(c).upper()]
                        
                        if r_name_c and r_price_c:
                            t2 = pd.DataFrame()
                            t2['Item Name'] = df[r_name_c[0]].astype(str) + " (Double/Right)"
                            t2['List Price'] = df[r_price_c[0]].apply(clean_price_value)
                            t2['Standard Discount'] = 0
                            t2['Category'] = sheet_name
                            t2['Unit'] = "Pc"
                            t2.dropna(subset=['List Price'], inplace=True)
                            all_data.append(t2)
                    except: pass

                if not temp.empty:
                    all_data.append(temp)
                    logs.append(f"✅ {sheet_name}: Found {len(temp)} items (Rate per **{detected_unit}**)")
            else:
                logs.append(f"❌ {sheet_name}: Columns missing.")

        except Exception as e:
            logs.append(f"⚠️ {sheet_name}: {str(e)}")

    if not all_data:
        return pd.DataFrame(), logs
        
    final_df = pd.concat(all_data, ignore_index=True)
    return final_df, logs

# --- 3. APP UI ---
catalog, logs = load_data_robust()

if 'cart' not in st.session_state: st.session_state['cart'] = []

# SIDEBAR
with st.sidebar:
    st.title("🔧 Config")
    
    with st.expander("Data Source Status"):
        for l in logs:
            if "✅" in l: st.markdown(l)
            else: st.error(l)

    # ADD ITEM
    st.header("1. Add Item")
    if not catalog.empty:
        cats = sorted(catalog['Category'].unique())
        sel_cat = st.selectbox("Category", cats)
        
        subset = catalog[catalog['Category'] == sel_cat]
        prods = sorted(subset['Item Name'].unique())
        sel_prod = st.selectbox("Product", prods)
        
        # Details
        row = subset[subset['Item Name'] == sel_prod].iloc[0]
        std_price = float(row['List Price'])
        std_disc = float(row['Standard Discount'])
        unit = row['Unit']
        
        # DISPLAY CLEAR RATE INFO
        st.info(f"**Rate:** ₹{std_price:,.2f} / **{unit}**\n\nDisc: {std_disc}%")
        
        c1, c2 = st.columns(2)
        qty = c1.number_input(f"Qty ({unit}s)", 1, 100000, 1)
        disc = c2.number_input("Discount %", 0.0, 100.0, std_disc)
        
        if st.button("Add to Quote", type="primary"):
            st.session_state['cart'].append({
                'Category': sel_cat,
                'Description': sel_prod,
                'Qty': qty,
                'Unit': unit,
                'List Price': std_price,
                'Discount': disc,
                'GST Rate': 18.0
            })
            st.success("Added!")
    
    # COLUMN MANAGER
    st.markdown("---")
    st.header("2. Table Columns")
    
    fixed_cols = ["No", "Description", "Qty", "Unit", "Total"]
    available_cols = ["Category", "List Price", "Discount %", "Net Rate", "GST %", "GST Amount"]
    default_cols = ["List Price", "Discount %", "Net Rate", "GST Amount"]
    
    selected_cols = st.multiselect("Select Extra Columns:", available_cols, default=default_cols)
    
    if st.button("Clear Cart"):
        st.session_state['cart'] = []
        st.rerun()

# MAIN PAGE
st.title("📄 Quotation System")

if not st.session_state['cart']:
    st.info("👈 Add items from the sidebar.")
else:
    data = []
    grand_total = 0
    
    for i, item in enumerate(st.session_state['cart']):
        lp = item['List Price']
        disc = item['Discount']
        qty = item['Qty']
        gst_rate = item.get('GST Rate', 18.0)
        
        # Financials
        net_rate = lp * (1 - disc/100)
        gst_amt = net_rate * (gst_rate/100)
        unit_total = net_rate + gst_amt
        line_total = unit_total * qty
        grand_total += line_total
        
        # Build Row - SAFE GET METHOD
        # using item.get('Unit', '-') prevents the crash for old items
        row = {
            "No": i+1,
            "Description": item['Description'],
            "Qty": qty,
            "Unit": item.get('Unit', '-'), 
            "Category": item['Category'],
            "List Price": f"{lp:,.2f}",
            "Discount %": f"{disc}%",
            "Net Rate": f"{net_rate:,.2f}",
            "GST %": f"{gst_rate}%",
            "GST Amount": f"{gst_amt:,.2f}",
            "Total": f"{line_total:,.2f}"
        }
        data.append(row)
        
    final_cols = ["No", "Description", "Qty", "Unit"] + selected_cols + ["Total"]
    
    df_disp = pd.DataFrame(data)
    for c in final_cols:
        if c not in df_disp.columns: df_disp[c] = ""
    df_disp = df_disp[final_cols]
    
    st.subheader("Draft Quotation")
    st.table(df_disp.set_index("No"))
    
    st.markdown(f"### Grand Total: ₹ {grand_total:,.2f}")
    
    st.markdown("""
    <button onclick="window.print()" style="
        background-color: #4CAF50; color: white; 
        padding: 12px 28px; border: none; border-radius: 5px; 
        font-size: 16px; cursor: pointer;">
        🖨️ Print / Save PDF
    </button>
    """, unsafe_allow_html=True)