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
    """ Decides if the Unit is 'Mtr', 'Pc', or 'Coil' """
    s_up = sheet_name.upper()
    c_up = col_name.upper()
    
    # 1. Explicit Column Name Wins
    if "MTR" in c_up or "METER" in c_up: return "Mtr"
    if "COIL" in c_up: return "Coil"
    if "PC" in c_up or "PIECE" in c_up: return "Pc"
        
    # 2. Category Heuristics
    if "GLAND" in s_up or "COSMOS" in s_up or "HMI" in s_up: return "Pc"
    if "CABLE" in s_up or "WIRE" in s_up or "ARM" in s_up: return "Mtr"
        
    return "Unit" # Fallback

# --- 2. DATA LOADER ---
@st.cache_data(show_spinner=True)
def load_data_robust():
    """ 
    Scans for ALL .xlsx files in current dir and 'data/' folder. 
    Uses Filename -> Main Category
    Uses Sheetname -> Sub Category
    """
    search_dirs = ['.', 'data']
    excel_files = []
    
    for d in search_dirs:
        if os.path.exists(d):
            files = [os.path.join(d, f) for f in os.listdir(d) 
                     if f.lower().endswith(".xlsx") and not f.startswith("~$")]
            excel_files.extend(files)
            
    if not excel_files:
        return pd.DataFrame(), ["❌ No .xlsx files found! Please add files to the folder."]

    all_data = []
    logs = []

    for file_path in excel_files:
        file_name = os.path.basename(file_path)
        main_category = os.path.splitext(file_name)[0]
        
        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                try:
                    df_raw = pd.read_excel(xls, sheet_name, header=None, nrows=60)
                    
                    # FIND HEADER
                    header_idx = -1
                    for i, row in df_raw.iterrows():
                        row_str = " ".join([str(x).upper() for x in row if pd.notna(x)])
                        if (("LP" in row_str or "PRICE" in row_str or "RATE" in row_str) and 
                            ("ITEM" in row_str or "DESC" in row_str or "CODE" in row_str or "SIZE" in row_str)):
                            header_idx = i
                            break
                    
                    if header_idx == -1:
                        for i, row in df_raw.iterrows():
                            if row.count() > 3:
                                header_idx = i
                                break

                    # READ DATA
                    df = pd.read_excel(xls, sheet_name, skiprows=header_idx)
                    df.columns = [str(c).strip() for c in df.columns]
                    
                    # IDENTIFY COLUMNS
                    name_col = None
                    price_col = None
                    disc_col = None
                    
                    # Price Priority
                    for col in df.columns:
                        c_up = col.upper()
                        if "MTR" in c_up or "METER" in c_up:
                            if any(x in c_up for x in ['LP', 'RATE', 'PRICE']):
                                price_col = col
                                break
                    if not price_col:
                        for col in df.columns:
                            c_up = col.upper()
                            if any(x in c_up for x in ['LP', 'PRICE', 'RATE', 'MRP', 'LIST', 'AMOUNT']):
                                if "AMOUNT" in c_up and any("RATE" in c.upper() for c in df.columns): continue
                                price_col = col
                                break

                    for col in df.columns:
                        c_up = col.upper()
                        if not name_col and any(x in c_up for x in ['DESC', 'ITEM', 'PARTICULARS', 'CODE', 'SIZE', 'TYPE', 'CABLE']):
                            name_col = col
                        if not disc_col and any(x in c_up for x in ['DISC', 'OFF']):
                            disc_col = col
                    
                    if not name_col: name_col = df.columns[0]
                    
                    # PROCESS
                    if name_col and price_col:
                        temp = pd.DataFrame()
                        temp['Item Name'] = df[name_col].astype(str)
                        temp['List Price'] = df[price_col].apply(clean_price_value)
                        
                        if disc_col:
                            temp['Standard Discount'] = pd.to_numeric(df[disc_col], errors='coerce').fillna(0)
                        else:
                            temp['Standard Discount'] = 0
                        
                        temp['Main Category'] = main_category
                        temp['Sub Category'] = sheet_name 
                        temp['Unit'] = detect_unit(sheet_name, price_col)
                        
                        temp.dropna(subset=['List Price'], inplace=True)
                        temp = temp[temp['List Price'] > 0]
                        
                        # HMI Special
                        if "HMI" in sheet_name.upper() and len(df.columns) > 8:
                            try:
                                r_name_c = [c for c in df.columns[6:] if "SIZE" in str(c).upper()]
                                r_price_c = [c for c in df.columns[6:] if "RATE" in str(c).upper()]
                                if r_name_c and r_price_c:
                                    t2 = pd.DataFrame()
                                    t2['Item Name'] = df[r_name_c[0]].astype(str) + " (Double/Right)"
                                    t2['List Price'] = df[r_price_c[0]].apply(clean_price_value)
                                    t2['Standard Discount'] = 0
                                    t2['Main Category'] = main_category
                                    t2['Sub Category'] = sheet_name
                                    t2['Unit'] = "Pc"
                                    t2.dropna(subset=['List Price'], inplace=True)
                                    all_data.append(t2)
                            except: pass

                        if not temp.empty:
                            all_data.append(temp)
                            logs.append(f"✅ Loaded: {file_name} -> {sheet_name}")
                    else:
                        pass 

                except Exception as e:
                    logs.append(f"⚠️ Error in {file_name} / {sheet_name}: {str(e)}")
        except Exception as e:
            logs.append(f"❌ Error reading file {file_name}: {str(e)}")

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
    if st.button("🔄 Refresh Data"):
        st.cache_data.clear()
        st.rerun()

    with st.expander("Data Logs"):
        if not logs: st.write("No files found.")
        for l in logs:
            if "✅" in l: st.markdown(l)
            else: st.error(l)

    st.header("1. Add Item")
    if not catalog.empty:
        # SELECTORS
        main_cats = sorted(catalog['Main Category'].unique())
        sel_main = st.selectbox("Main Category", main_cats)
        
        subset_main = catalog[catalog['Main Category'] == sel_main]
        sub_cats = sorted(subset_main['Sub Category'].unique())
        sel_sub = st.selectbox("Sub Category", sub_cats)
        
        subset_final = subset_main[subset_main['Sub Category'] == sel_sub]
        prods = sorted(subset_final['Item Name'].unique())
        sel_prod = st.selectbox("Product", prods)
        
        # DETAILS
        row = subset_final[subset_final['Item Name'] == sel_prod].iloc[0]
        std_price = float(row['List Price'])
        std_disc = float(row['Standard Discount'])
        unit = row['Unit']
        
        st.info(f"**Rate:** ₹{std_price:,.2f} / **{unit}**\n\nStd Disc: {std_disc}%")
        
        # INPUTS
        c1, c2 = st.columns(2)
        qty = c1.number_input(f"Qty ({unit}s)", 1, 100000, 1)
        disc = c2.number_input("Discount %", 0.0, 100.0, std_disc)
        make_val = st.text_input("Make / Manufacturer", value="", placeholder="e.g. Polycab")
        
        if st.button("Add to Quote", type="primary"):
            st.session_state['cart'].append({
                'Main Category': sel_main,
                'Sub Category': sel_sub,
                'Description': sel_prod,
                'Make': make_val, 
                'Qty': qty,
                'Unit': unit,
                'List Price': std_price,
                'Discount': disc,
                'GST Rate': 18.0
            })
            st.success("Added!")

    st.markdown("---")
    st.header("2. Table Columns")
    available_cols = ["Main Category", "Sub Category", "List Price", "Discount %", "Net Rate", "GST %", "GST Amount"]
    default_cols = ["Make", "List Price", "Discount %", "Net Rate", "GST Amount"]
    selected_cols = st.multiselect("Select Extra Columns:", available_cols, default=default_cols)
    
    if st.button("🗑️ Clear Entire Cart"):
        st.session_state['cart'] = []
        st.rerun()

# MAIN PAGE
st.title("📄 Quotation System")

if not st.session_state['cart']:
    st.info("👈 Add items from the sidebar to start.")
else:
    # --- SECTION A: REVIEW & EDIT (DELETE ITEMS) ---
    st.subheader("1. Review & Edit Items")
    st.markdown("Use the trash icon to remove a specific item.")
    
    # Header Row
    h1, h2, h3, h4, h5, h6 = st.columns([0.5, 4, 2, 1.5, 2, 1])
    h1.markdown("**#**")
    h2.markdown("**Description**")
    h3.markdown("**Make**")
    h4.markdown("**Qty**")
    h5.markdown("**Total**")
    h6.markdown("**Action**")
    st.markdown("---")

    # Item Rows
    for i, item in enumerate(st.session_state['cart']):
        # Calculate totals for preview
        lp = item['List Price']
        disc_val = item['Discount']
        q = item['Qty']
        g_rate = item.get('GST Rate', 18.0)
        net = lp * (1 - disc_val/100)
        gst = net * (g_rate/100)
        tot = (net + gst) * q
        
        c1, c2, c3, c4, c5, c6 = st.columns([0.5, 4, 2, 1.5, 2, 1])
        c1.write(f"{i+1}")
        c2.write(item['Description'])
        c3.write(item.get('Make', '-'))
        c4.write(f"{q}")
        c5.write(f"₹ {tot:,.0f}")
        
        # DELETE BUTTON
        # key=f"del_{i}" ensures every button is unique
        if c6.button("🗑️", key=f"del_{i}", help="Remove this item"):
            st.session_state['cart'].pop(i)
            st.rerun()

    st.markdown("---")
    
    # --- SECTION B: FINAL TABLE (FOR PRINTING) ---
    st.subheader("2. Final Draft Quotation")
    
    data = []
    grand_total = 0
    
    for i, item in enumerate(st.session_state['cart']):
        lp = item['List Price']
        disc_val = item['Discount']
        qty = item['Qty']
        gst_rate = item.get('GST Rate', 18.0)
        
        net_rate = lp * (1 - disc_val/100)
        gst_amt = net_rate * (gst_rate/100)
        unit_total = net_rate + gst_amt
        line_total = unit_total * qty
        grand_total += line_total
        
        row = {
            "No": i+1,
            "Description": item['Description'],
            "Make": item.get('Make', ''),
            "Qty": qty,
            "Unit": item.get('Unit', '-'), 
            "Main Category": item['Main Category'],
            "Sub Category": item['Sub Category'],
            "List Price": f"{lp:,.2f}",
            "Discount %": f"{disc_val}%",
            "Net Rate": f"{net_rate:,.2f}",
            "GST %": f"{gst_rate}%",
            "GST Amount": f"{gst_amt:,.2f}",
            "Total": f"{line_total:,.2f}"
        }
        data.append(row)
        
    # Column ordering
    final_cols = ["No", "Description"] 
    if "Make" in selected_cols: final_cols.append("Make")
    final_cols = final_cols + ["Qty", "Unit"]
    for c in selected_cols:
        if c not in final_cols: final_cols.append(c)
    final_cols.append("Total")
    
    df_disp = pd.DataFrame(data)
    for c in final_cols:
        if c not in df_disp.columns: df_disp[c] = ""
    df_disp = df_disp[final_cols]
    
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