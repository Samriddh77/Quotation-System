import streamlit as st
import pandas as pd
import os
import re

# --- PAGE CONFIG ---
st.set_page_config(page_title="Quotation Generator", layout="wide")

# --- 1. HELPERS ---
def clean_price_value(val):
    if pd.isna(val): return 0.0
    s = str(val).strip()
    s_clean = re.sub(r'[^\d.]', '', s)
    try: return float(s_clean)
    except: return 0.0

def clean_coil_len(val):
    if pd.isna(val): return 0.0
    s = str(val).strip()
    s_clean = re.sub(r'[^\d.]', '', s)
    try: return float(s_clean)
    except: return 0.0

def detect_uom(sheet_name, price_col_name):
    s_up = str(sheet_name).upper()
    c_up = str(price_col_name).upper()
    if "MTR" in c_up or "METER" in c_up: return "Mtr"
    if "PC" in c_up or "PIECE" in c_up: return "Pc"
    if "GLAND" in s_up or "HMI" in s_up or "COSMOS" in s_up: return "Pc"
    return "Mtr" 

# --- 2. DATA LOADER ---
@st.cache_data(show_spinner=True)
def load_data_from_files():
    search_dirs = ['.', 'data'] 
    all_dfs = []
    logs = []

    excel_files = []
    for d in search_dirs:
        if os.path.exists(d):
            files = [os.path.join(d, f) for f in os.listdir(d) 
                     if f.lower().endswith(".xlsx") and not f.startswith("~$")]
            excel_files.extend(files)
    
    if not excel_files: return pd.DataFrame(), ["❌ No .xlsx files found!"]

    for file_path in excel_files:
        filename = os.path.basename(file_path)
        main_cat = os.path.splitext(filename)[0] 
        
        try:
            xls = pd.ExcelFile(file_path)
            for sheet in xls.sheet_names:
                try:
                    df_raw = pd.read_excel(xls, sheet, header=None, nrows=30)
                    header_idx = -1
                    
                    for i, row in df_raw.iterrows():
                        r_str = " ".join([str(x).upper() for x in row if pd.notna(x)])
                        if (("LP" in r_str or "PRICE" in r_str or "RATE" in r_str) and 
                            ("ITEM" in r_str or "DESC" in r_str or "SIZE" in r_str or "CODE" in r_str)):
                            header_idx = i
                            break
                    
                    if header_idx == -1: continue

                    df = pd.read_excel(xls, sheet, skiprows=header_idx)
                    df.columns = [str(c).strip() for c in df.columns]
                    
                    name_col = None
                    price_col = None
                    disc_col = None
                    coil_col = None
                    
                    price_col = next((c for c in df.columns if "PER MTR" in c.upper() or "PER METER" in c.upper()), None)
                    if not price_col:
                        for k in ["LP", "LIST PRICE", "RATE", "PRICE", "MRP"]:
                            match = next((c for c in df.columns if k in c.upper()), None)
                            if match and "AMOUNT" not in match.upper(): price_col = match; break
                    
                    possible_names = ["ITEM DESCRIPTION", "DESCRIPTION", "PARTICULARS", "SIZE", "CODE", "ITEM"]
                    for k in possible_names:
                        match = next((c for c in df.columns if k in c.upper()), None)
                        if match: name_col = match; break
                    if not name_col: name_col = df.columns[0]

                    disc_col = next((c for c in df.columns if "DISC" in c.upper() or "OFF" in c.upper()), None)
                    coil_col = next((c for c in df.columns if "COIL" in c.upper() and ("LEN" in c.upper() or "MTR" in c.upper()) and "QTY" not in c.upper()), None)

                    if price_col:
                        clean_df = pd.DataFrame()
                        clean_df['Description'] = df[name_col].astype(str)
                        clean_df['List Price'] = df[price_col].apply(clean_price_value)
                        
                        if disc_col: clean_df['Standard Discount'] = pd.to_numeric(df[disc_col], errors='coerce').fillna(0)
                        else: clean_df['Standard Discount'] = 0
                        
                        if coil_col: clean_df['Coil Length'] = df[coil_col].apply(clean_coil_len)
                        else: clean_df['Coil Length'] = 0.0

                        clean_df['Main Category'] = main_cat
                        clean_df['Sub Category'] = sheet
                        
                        if "UOM" in [c.upper() for c in df.columns]:
                            uom_c = next(c for c in df.columns if "UOM" in c.upper())
                            clean_df['UOM'] = df[uom_c].astype(str)
                        else:
                            clean_df['UOM'] = detect_uom(sheet, price_col)

                        clean_df = clean_df[clean_df['List Price'] > 0]
                        clean_df = clean_df[clean_df['Description'] != 'nan']
                        
                        all_dfs.append(clean_df)
                        logs.append(f"✅ Loaded {main_cat} -> {sheet}")
                except: pass

        except Exception as e: logs.append(f"❌ Error {filename}: {e}")

    if not all_dfs: return pd.DataFrame(), logs
    return pd.concat(all_dfs, ignore_index=True), logs

# --- 3. APP UI ---
catalog, logs = load_data_from_files()

if 'cart' not in st.session_state: st.session_state['cart'] = []

# SIDEBAR
with st.sidebar:
    st.title("🔧 Config")
    if st.button("🔄 Refresh Data"):
        st.cache_data.clear()
        st.rerun()
    
    with st.expander("System Logs"):
        for l in logs: st.write(l)

    st.header("1. Add Item")
    
    if not catalog.empty:
        main_cats = sorted(catalog['Main Category'].unique())
        sel_main = st.selectbox("1. Category", main_cats)
        
        subset_main = catalog[catalog['Main Category'] == sel_main]
        sub_cats = sorted(subset_main['Sub Category'].unique())
        sel_sub = st.selectbox("2. Type/Brand", sub_cats)
        
        subset_final = subset_main[subset_main['Sub Category'] == sel_sub]
        prods = sorted(subset_final['Description'].unique())
        sel_prod = st.selectbox("3. Product", prods)
        
        row = subset_final[subset_final['Description'] == sel_prod].iloc[0]
        std_price = float(row['List Price'])
        std_disc = float(row['Standard Discount'])
        uom = row['UOM']
        coil_len = float(row['Coil Length'])
        
        st.info(f"Rate: ₹{std_price:,.2f} / {uom}")
        
        calc_qty = 0
        disp_unit = uom
        
        if coil_len > 0 and "M" in uom.upper():
            st.caption(f"Standard Coil: {int(coil_len)} Mtr")
            mode = st.radio("Input:", ["Coils", "Meters"], horizontal=True, label_visibility="collapsed")
            
            if mode == "Coils":
                q_in = st.number_input("No. of Coils", 1, 500, 1)
                calc_qty = q_in * coil_len
                st.write(f"= **{calc_qty} Mtr**")
                disp_unit = f"{uom} ({q_in} Coils)"
            else:
                q_in = st.number_input("Total Meters", 1, 10000, 100)
                calc_qty = q_in
                st.caption(f"Approx {q_in/coil_len:.1f} Coils")
                disp_unit = uom
        else:
            q_in = st.number_input(f"Qty ({uom})", 1, 10000, 1)
            calc_qty = q_in
            disp_unit = uom
            
        c1, c2 = st.columns(2)
        disc = c1.number_input("Disc %", 0.0, 100.0, std_disc)
        make = c2.text_input("Make", value="", placeholder="Brand")
        
        # --- NEW REMARK FIELD ---
        remark = st.text_input("Remark", value="", placeholder="e.g. Urgent / Red Color")
        
        if st.button("Add Item", type="primary"):
            st.session_state['cart'].append({
                'Main Category': sel_main,
                'Sub Category': sel_sub,
                'Description': sel_prod,
                'Make': make,
                'Remark': remark, # <--- Saved here
                'Qty': calc_qty,
                'Display Unit': disp_unit,
                'List Price': std_price,
                'Discount': disc,
                'GST Rate': 18.0
            })
            st.success("Added")

    st.markdown("---")
    st.header("2. Columns")
    # Added "Remark" to available options
    avail_cols = ["Make", "Remark", "Main Category", "Sub Category", "List Price", "Discount %", "Net Rate", "GST %", "GST Amount"]
    def_cols = ["Make", "Remark", "List Price", "Discount %", "Net Rate", "GST Amount"]
    
    sel_cols = st.multiselect("Select Cols", avail_cols, default=def_cols)
    
    if st.button("🗑️ Clear Cart"):
        st.session_state['cart'] = []
        st.rerun()

# --- MAIN PAGE ---
st.title("📄 Quotation System")

if not st.session_state['cart']:
    st.info("👈 Add items from the sidebar.")
else:
    # 1. REVIEW SECTION
    st.subheader("1. Review Items")
    h1, h2, h3, h4, h5, h6 = st.columns([0.5, 4, 1.5, 1.5, 2, 0.5])
    h1.write("#"); h2.write("Item"); h3.write("Make"); h4.write("Qty"); h5.write("Total"); 
    st.divider()
    
    for i, item in enumerate(st.session_state['cart']):
        lp = item['List Price']
        d = item['Discount']
        q = item['Qty']
        g = item.get('GST Rate', 18.0)
        
        net = lp * (1 - d/100)
        gst_amt = net * (g/100)
        tot = (net + gst_amt) * q
        
        c1, c2, c3, c4, c5, c6 = st.columns([0.5, 4, 1.5, 1.5, 2, 0.5])
        c1.write(f"{i+1}")
        c2.write(item['Description'])
        c3.write(item.get('Make', '-'))
        c4.write(f"{item['Display Unit']}") 
        c5.write(f"₹ {tot:,.0f}")
        if c6.button("🗑️", key=f"d{i}"):
            st.session_state['cart'].pop(i)
            st.rerun()

    st.divider()
    
    # 2. FINAL TABLE GENERATION
    st.subheader("2. Final Draft")
    
    final_data = []
    
    sum_taxable = 0
    sum_gst = 0
    sum_grand = 0
    
    for i, item in enumerate(st.session_state['cart']):
        lp = item['List Price']
        d = item['Discount']
        q = item['Qty']
        g = item.get('GST Rate', 18.0)
        
        unit_net = lp * (1 - d/100)
        line_taxable = unit_net * q
        unit_gst = unit_net * (g/100)
        line_gst = unit_gst * q
        line_total = line_taxable + line_gst
        
        sum_taxable += line_taxable
        sum_gst += line_gst
        sum_grand += line_total
        
        row = {
            "No": str(i+1),
            "Description": item['Description'],
            "Make": item.get('Make', ''),
            "Remark": item.get('Remark', ''), # <--- Added here
            "Qty": f"{q:,.2f}",
            "Unit": item['Display Unit'],
            "Main Category": item['Main Category'],
            "Sub Category": item['Sub Category'],
            "List Price": f"{lp:,.2f}",
            "Discount %": f"{d}%",
            "Net Rate": f"{unit_net:,.2f}",
            "GST %": f"{g}%",
            "GST Amount": f"{line_gst:,.2f}",
            "Total": f"{line_total:,.2f}"
        }
        final_data.append(row)
        
    # Spacer Row
    final_data.append({"Description": "", "Total": ""}) 
    
    # Summary Rows
    final_data.append({
        "No": "", "Description": "TOTAL (BEFORE GST)", 
        "Total": f"₹ {sum_taxable:,.2f}"
    })
    final_data.append({
        "No": "", "Description": "TOTAL GST AMOUNT", 
        "Total": f"₹ {sum_gst:,.2f}"
    })
    final_data.append({
        "No": "", "Description": "GRAND TOTAL (INCL. GST)", 
        "Total": f"₹ {sum_grand:,.2f}"
    })
        
    df_out = pd.DataFrame(final_data)
    
    # Build Columns
    show_cols = ["No", "Description"]
    
    # Logic to insert Make and Remark if selected
    if "Make" in sel_cols: show_cols.append("Make")
    if "Remark" in sel_cols: show_cols.append("Remark")
    
    show_cols += ["Qty", "Unit"]
    
    for c in sel_cols:
        if c not in show_cols: show_cols.append(c)
    
    if "Total" not in show_cols: show_cols.append("Total")
    
    # Fill NaN
    for c in show_cols: 
        if c not in df_out.columns: df_out[c] = ""
        
    st.table(df_out[show_cols].set_index("No"))
    
    st.markdown("""
    <button onclick="window.print()" style="
        background-color: #4CAF50; color: white; 
        padding: 12px 28px; border: none; border-radius: 5px; 
        font-size: 16px; cursor: pointer;">
        🖨️ Print / Save PDF
    </button>
    """, unsafe_allow_html=True)