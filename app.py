import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime
from io import BytesIO

try:
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
except ImportError:
    st.error("❌ Library 'python-docx' is missing. Please run: pip install python-docx")
    st.stop()

# --- 1. SETUP PATHS ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, 'templates')
DATA_DIR = os.path.join(BASE_DIR, 'data')

FIRM_MAPPING = {
    "Electro World": "Electro_Template.docx",
    "Abhinav Enterprises": "Abhinav_Template.docx",
    "Shree Creative Marketing": "Shree_Template.docx"
}

# --- 2. PRESERVE ORIGINAL TEMPLATES ---
# We only create folders. We DO NOT generate dummy files if you have uploaded your own.
if not os.path.exists(DATA_DIR): os.makedirs(DATA_DIR)
if not os.path.exists(TEMPLATE_DIR): os.makedirs(TEMPLATE_DIR)

# --- CONFIGURATION ---
st.set_page_config(page_title="Quotation Generator", layout="wide")

FIRM_DEFAULTS = {
    "Electro World": {
        "price": "Nett", "gst": "Extra @ 18%", "delivery": "Ex Stock / 7 Days",
        "freight": "Ex Our Godown", "payment": "100% Against Delivery",
        "validity": "7 Days", "guarantee": "12 months"
    },
    "Abhinav Enterprises": {
        "price": "Nett", "gst": "18% Extra", "delivery": "Within 2-3 Weeks",
        "freight": "P&F: NIL | F.O.R: Ex Indore", "payment": "30% Advance",
        "validity": "3 Days", "guarantee": "12 months"
    },
    "Shree Creative Marketing": {
        "price": "Nett", "gst": "Extra @ 18%", "delivery": "Ex Stock",
        "freight": "To Pay Basis", "payment": "Immediate",
        "validity": "5 Days", "guarantee": "Standard Mfg Warranty"
    }
}

# --- STATE MANAGEMENT ---
def update_defaults():
    firm = st.session_state.get('firm_selector', "Electro World")
    defaults = FIRM_DEFAULTS.get(firm, FIRM_DEFAULTS["Electro World"])
    st.session_state['p_term'] = defaults['price']
    st.session_state['g_term'] = defaults['gst']
    st.session_state['d_term'] = defaults['delivery']
    st.session_state['f_term'] = defaults['freight']
    st.session_state['pay_term'] = defaults['payment']
    st.session_state['val_term'] = defaults['validity']
    st.session_state['guar_term'] = defaults['guarantee']

if 'firm_selector' not in st.session_state: st.session_state['firm_selector'] = "Electro World"
if 'p_term' not in st.session_state: update_defaults()

# --- HELPERS ---
def clean_price_value(val):
    if pd.isna(val): return 0.0
    s = str(val).strip()
    s_clean = re.sub(r'[^\d.]', '', s)
    try: return float(s_clean)
    except: return 0.0

def detect_uom(sheet_name, price_col_name):
    c_up = str(price_col_name).upper()
    if "MTR" in c_up or "METER" in c_up: return "Mtr"
    if "PC" in c_up or "PIECE" in c_up: return "Pc"
    return "Mtr" 

# --- WORD GENERATOR UTILS ---
def set_table_borders(table):
    """ Force proper grid borders using XML """
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblBorders = tblPr.first_child_found_in("w:tblBorders")
    if tblBorders is None:
        tblBorders = OxmlElement('w:tblBorders')
        tblPr.append(tblBorders)
    
    for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        tblBorders.append(border)

def safe_replace_text(doc, replacements):
    """ Replaces text while PRESERVING formatting (Bold, Fonts, Colors) """
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                # We do a naive replace first to catch simple cases, 
                # but for complex formatting we should iterate runs if needed.
                # However, python-docx simple replace often kills style.
                # This approach tries to preserve the run structure.
                for run in paragraph.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, value)
    
    # Also replace in tables (Header section etc)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key, value in replacements.items():
                        if key in paragraph.text:
                            for run in paragraph.runs:
                                if key in run.text:
                                    run.text = run.text.replace(key, value)

def fill_template_docx(template_path, client_data, cart_items, terms, visible_cols):
    doc = Document(template_path)
    
    replacements = {
        '{{REF_NO}}': str(client_data.get('ref_no', '')),
        '{{DATE}}': datetime.now().strftime('%d-%b-%Y'),
        '{{CLIENT_NAME}}': str(client_data.get('client_name', '')),
        '{{CLIENT_ADDRESS}}': str(client_data.get('client_address', '')),
        '{{SUBJECT}}': str(client_data.get('subject', '')),
        '{{PRICE_TERM}}': str(terms.get('price_term', '')),
        '{{GST_TERM}}': str(terms.get('gst_term', '')),
        '{{DELIVERY_TERM}}': str(terms.get('delivery_term', '')),
        '{{FREIGHT_TERM}}': str(terms.get('freight_term', '')),
        '{{PAYMENT_TERM}}': str(terms.get('payment_term', '')),
        '{{VALIDITY_TERM}}': str(terms.get('validity_term', '')),
        '{{GUARANTEE_TERM}}': str(terms.get('guarantee_term', '')),
        '{{GURANTEE_TERM}}': str(terms.get('guarantee_term', '')), # Handle typo in template
    }

    # 1. Perform Text Replacements (Preserving Style)
    safe_replace_text(doc, replacements)

    # 2. Find Table Placeholder
    target_paragraph = None
    for paragraph in doc.paragraphs:
        if '{{TABLE_HERE}}' in paragraph.text:
            target_paragraph = paragraph
            break
            
    if target_paragraph:
        target_paragraph.text = "" # Clear the placeholder tag
        
        # Define Column Config
        col_defs = {
            "S.No.": {"width": Cm(1.2), "align": WD_ALIGN_PARAGRAPH.CENTER},
            "Sub Category": {"width": Cm(3.0), "align": WD_ALIGN_PARAGRAPH.LEFT},
            "Item Description": {"width": Cm(6.5), "align": WD_ALIGN_PARAGRAPH.LEFT},
            "Make": {"width": Cm(2.5), "align": WD_ALIGN_PARAGRAPH.CENTER},
            "Qty": {"width": Cm(1.5), "align": WD_ALIGN_PARAGRAPH.RIGHT},
            "Unit": {"width": Cm(1.5), "align": WD_ALIGN_PARAGRAPH.CENTER},
            "Rate": {"width": Cm(2.5), "align": WD_ALIGN_PARAGRAPH.RIGHT},
            "Amount": {"width": Cm(3.0), "align": WD_ALIGN_PARAGRAPH.RIGHT}
        }
        
        active_headers = [h for h in visible_cols if h in col_defs]
        
        # Create Table
        table = doc.add_table(rows=1, cols=len(active_headers))
        table.autofit = False
        set_table_borders(table)
        
        # Header Row
        for i, header in enumerate(active_headers):
            cell = table.rows[0].cells[i]
            cell.text = header
            cell.width = col_defs[header]["width"]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            # Make Bold
            if p.runs: p.runs[0].bold = True
            else: p.add_run(header).bold = True

        # Data Rows
        total_amt = 0
        for i, item in enumerate(cart_items):
            row_cells = table.add_row().cells
            
            lp = item['List Price']
            disc = item['Discount']
            qty = item['Qty']
            net_rate = lp * (1 - disc/100)
            line_total = net_rate * qty
            total_amt += line_total
            
            make_str = item.get('Make', '').strip()
            if make_str: make_str = f"{make_str} Make"
            
            data_map = {
                "S.No.": str(i+1),
                "Sub Category": item.get('Sub Category', ''),
                "Item Description": item['Description'],
                "Make": make_str,
                "Qty": f"{qty:,.2f}",
                "Unit": item['Display Unit'].split()[0],
                "Rate": f"{net_rate:,.2f}",
                "Amount": f"{line_total:,.2f}"
            }
            
            for idx, header in enumerate(active_headers):
                cell = row_cells[idx]
                cell.text = data_map.get(header, "")
                cell.width = col_defs[header]["width"]
                cell.paragraphs[0].alignment = col_defs[header]["align"]

        # Grand Total Row
        if "Amount" in active_headers:
            row = table.add_row().cells
            amt_idx = active_headers.index("Amount")
            label_idx = amt_idx - 1 if amt_idx > 0 else 0
            
            # Label
            label_cell = row[label_idx]
            label_cell.text = "Grand Total (Excl. GST)"
            label_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if label_cell.paragraphs[0].runs: label_cell.paragraphs[0].runs[0].bold = True
            else: label_cell.paragraphs[0].add_run("Grand Total (Excl. GST)").bold = True

            # Value
            val_cell = row[amt_idx]
            val_cell.text = f"{total_amt:,.2f}"
            val_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if val_cell.paragraphs[0].runs: val_cell.paragraphs[0].runs[0].bold = True
            else: val_cell.paragraphs[0].add_run(f"{total_amt:,.2f}").bold = True

        target_paragraph._p.addnext(table._tbl)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- DATA LOADER ---
@st.cache_data(show_spinner=True)
def load_data_from_files():
    all_dfs = []
    if not os.path.exists(DATA_DIR): return pd.DataFrame(), ["❌ Data folder missing"]
    
    excel_files = [os.path.join(DATA_DIR, f) for f in os.listdir(DATA_DIR) 
                   if f.lower().endswith(".xlsx") and not f.startswith("~$")]
    
    if not excel_files: return pd.DataFrame(), ["⚠️ No .xlsx files found."]

    for file_path in excel_files:
        filename = os.path.basename(file_path)
        try:
            xls = pd.ExcelFile(file_path)
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet)
                df.columns = [str(c).strip() for c in df.columns]
                
                name_col = next((c for c in df.columns if any(k in c.upper() for k in ["DESC", "ITEM", "PARTICULARS"])), None)
                price_col = next((c for c in df.columns if any(k in c.upper() for k in ["PRICE", "RATE", "LP"]) and "AMOUNT" not in c.upper()), None)
                
                if name_col and price_col:
                    clean_df = pd.DataFrame()
                    clean_df['Description'] = df[name_col].astype(str)
                    clean_df['List Price'] = df[price_col].apply(clean_price_value)
                    clean_df['Discount'] = 0.0
                    clean_df['Main Category'] = os.path.splitext(filename)[0]
                    clean_df['Sub Category'] = sheet
                    clean_df['UOM'] = detect_uom(sheet, price_col)
                    clean_df['Display Unit'] = clean_df['UOM']
                    all_dfs.append(clean_df[clean_df['List Price'] > 0])
        except: pass

    if not all_dfs: return pd.DataFrame(), ["⚠️ Error reading Excel."]
    return pd.concat(all_dfs, ignore_index=True), []

# --- APP UI ---
catalog, logs = load_data_from_files()
if 'cart' not in st.session_state: st.session_state['cart'] = []

with st.sidebar:
    st.title("🔧 Config")
    if st.button("🔄 Refresh Data"): st.cache_data.clear(); st.rerun()
    st.markdown("---")
    st.header("1. Add Item")
    
    if not catalog.empty:
        cats = sorted(catalog['Main Category'].unique())
        sel_cat = st.selectbox("Category", cats)
        sub_cats = sorted(catalog[catalog['Main Category'] == sel_cat]['Sub Category'].unique())
        sel_sub = st.selectbox("Sub Category", sub_cats)
        
        subset = catalog[(catalog['Main Category'] == sel_cat) & (catalog['Sub Category'] == sel_sub)]
        prods = sorted(subset['Description'].unique())
        sel_prod = st.selectbox("Product", prods)
        
        row = subset[subset['Description'] == sel_prod].iloc[0]
        std_price = row['List Price']
        uom = row['UOM']
        
        st.info(f"Price: {std_price} / {uom}")
        calc_qty = st.number_input(f"Qty ({uom})", 1, 10000)
        c1, c2 = st.columns(2)
        disc = c1.number_input("Disc %", 0.0, 100.0, 0.0)
        make = c2.text_input("Make (e.g. HMI)")
        
        if st.button("Add"):
            st.session_state['cart'].append({
                'Main Category': sel_cat, 'Sub Category': sel_sub,
                'Description': sel_prod, 'Make': make,
                'Qty': calc_qty, 'Display Unit': uom,
                'List Price': std_price, 'Discount': disc
            })
            st.success("Added")
    else:
        st.warning("Please add an Excel file to the 'data' folder.")
            
    if st.button("Clear Cart"): st.session_state['cart'] = []; st.rerun()

st.title("📄 Quotation System")

if not st.session_state['cart']:
    st.info("Add items from the sidebar to start.")
else:
    st.subheader("1. Item List")
    col_config = [0.5, 3.5, 1.5, 1.5, 1.2, 1.5, 0.5]
    h1, h2, h3, h4, h5, h6, h7 = st.columns(col_config)
    h1.write("#"); h2.write("Desc"); h3.write("Make"); h4.write("Qty"); h5.write("Unit"); h6.write("Total"); 
    st.divider()

    grand_tot = 0
    for i, item in enumerate(st.session_state['cart']):
        net = item['List Price'] * (1 - item['Discount']/100)
        tot = net * item['Qty']
        grand_tot += tot
        c1, c2, c3, c4, c5, c6, c7 = st.columns(col_config)
        c1.write(f"{i+1}"); c2.write(item['Description'])
        make_display = f"{item['Make']} Make" if item['Make'] else ""
        c3.write(make_display); c4.write(f"{item['Qty']:,.2f}"); c5.write(item['Display Unit']); c6.write(f"₹ {tot:,.0f}")
        if c7.button("x", key=f"d{i}"):
            st.session_state['cart'].pop(i)
            st.rerun()

    st.divider()
    st.write(f"**Total (Excl. GST): ₹ {grand_tot:,.2f}**")
    st.markdown("---")

    st.subheader("2. Generate Quotation")
    selected_firm = st.selectbox("Select Template (Firm)", list(FIRM_MAPPING.keys()), key="firm_selector", on_change=update_defaults)
    
    # --- TABLE SETTINGS ---
    st.markdown("##### Table Settings")
    all_cols = ["S.No.", "Sub Category", "Item Description", "Make", "Qty", "Unit", "Rate", "Amount"]
    visible_cols = st.multiselect("Columns to include", all_cols, default=all_cols)
    
    c1, c2 = st.columns(2)
    with c1:
        client_name = st.text_input("Client Name", placeholder="M/s Client Name")
        ref_no = st.text_input("Ref No", value=f"QTN/{datetime.now().strftime('%y%m%d')}/001")
        subject = st.text_input("Subject", value="OFFER FOR ELECTRICAL GOODS")
    with c2:
        client_address = st.text_area("Client Address", placeholder="Address Line 1\nCity, State")
    
    st.markdown("#### Terms & Conditions")
    tc1, tc2, tc3 = st.columns(3)
    p_term = tc1.text_input("Price Term", key='p_term')
    g_term = tc2.text_input("GST Term", key='g_term')
    d_term = tc3.text_input("Delivery", key='d_term')
    tc4, tc5, tc6 = st.columns(3)
    f_term = tc4.text_input("Freight", key='f_term')
    pay_term = tc5.text_input("Payment", key='pay_term')
    val_term = tc6.text_input("Validity", key='val_term')
    guarantee = st.text_input("Guarantee", key='guar_term')

    if st.button("📥 Download Word Document", type="primary"):
        if not client_name:
            st.error("Please enter Client Name.")
        else:
            client_data = {"client_name": client_name, "client_address": client_address, "ref_no": ref_no, "subject": subject}
            terms = {"price_term": p_term, "gst_term": g_term, "delivery_term": d_term, "freight_term": f_term, "payment_term": pay_term, "validity_term": val_term, "guarantee_term": guarantee}
            
            template_path = os.path.join(TEMPLATE_DIR, FIRM_MAPPING[selected_firm])
            
            if not os.path.exists(template_path):
                 st.error(f"❌ Template not found: {template_path}. Please upload it to the 'templates' folder.")
            else:
                docx_file = fill_template_docx(template_path, client_data, st.session_state['cart'], terms, visible_cols)
                st.download_button(
                    label=f"Download Quote for {selected_firm}",
                    data=docx_file,
                    file_name=f"Quote_{selected_firm.replace(' ','_')}_{client_name[:5]}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )