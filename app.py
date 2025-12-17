import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime
from io import BytesIO

try:
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    st.error("❌ Library 'python-docx' is missing. Please run: pip install python-docx")
    st.stop()

# --- 1. SELF-HEALING SETUP (Fixes Missing Templates & Folders) ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, 'templates')
DATA_DIR = os.path.join(BASE_DIR, 'data')

FIRM_MAPPING = {
    "Electro World": "Electro_Template.docx",
    "Abhinav Enterprises": "Abhinav_Template.docx",
    "Shree Creative Marketing": "Shree_Template.docx"
}

def check_and_create_system():
    # 1. Create Data Folder if missing
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)
    
    # 2. Create Template Folder if missing
    if not os.path.exists(TEMPLATE_DIR):
        os.makedirs(TEMPLATE_DIR)

    # 3. Create missing .docx templates with placeholders
    for firm, filename in FIRM_MAPPING.items():
        file_path = os.path.join(TEMPLATE_DIR, filename)
        if not os.path.exists(file_path):
            doc = Document()
            doc.add_heading(f'QUOTATION - {firm}', 0)
            doc.add_paragraph(f"Ref No: {{{{REF_NO}}}}")
            doc.add_paragraph(f"Date: {{{{DATE}}}}")
            doc.add_paragraph("To,\n{{CLIENT_NAME}}\n{{CLIENT_ADDRESS}}")
            doc.add_paragraph("Subject: {{SUBJECT}}")
            doc.add_paragraph("Dear Sir/Ma'am,\nPlease find our offer below:")
            
            # THE CRITICAL PLACEHOLDER FOR THE TABLE
            doc.add_paragraph("{{TABLE_HERE}}")
            
            doc.add_heading('Terms and Conditions:', level=2)
            doc.add_paragraph("Price: {{PRICE_TERM}}")
            doc.add_paragraph("GST: {{GST_TERM}}")
            doc.add_paragraph("Delivery: {{DELIVERY_TERM}}")
            doc.add_paragraph("Freight: {{FREIGHT_TERM}}")
            doc.add_paragraph("Payment: {{PAYMENT_TERM}}")
            doc.add_paragraph("Validity: {{VALIDITY_TERM}}")
            doc.add_paragraph("Guarantee: {{GUARANTEE_TERM}}")
            
            doc.save(file_path)

# Run setup immediately on app start
check_and_create_system()

# --- CONFIGURATION ---
st.set_page_config(page_title="Quotation Generator", layout="wide")

# DEFAULT TERMS
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

if 'firm_selector' not in st.session_state:
    st.session_state['firm_selector'] = "Electro World"

if 'p_term' not in st.session_state:
    update_defaults()

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

# --- WORD GENERATOR (FIXED) ---
def replace_text_in_paragraph(paragraph, key, value):
    if key in paragraph.text:
        paragraph.text = paragraph.text.replace(key, value)

def fill_template_docx(template_path, client_data, cart_items, terms):
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
    }

    # Replace in Paragraphs
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            replace_text_in_paragraph(paragraph, key, value)

    # Replace in Tables (if any exist in template)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key, value in replacements.items():
                        replace_text_in_paragraph(paragraph, key, value)

    # Insert Item Table
    target_paragraph = None
    for paragraph in doc.paragraphs:
        if '{{TABLE_HERE}}' in paragraph.text:
            target_paragraph = paragraph
            break
            
    if target_paragraph:
        target_paragraph.text = "" 
        
        headers = ["S.No.", "Item Description", "Qty", "Unit", "Rate", "Amount"]
        table = doc.add_table(rows=1, cols=len(headers))
        
        # --- FIX: SAFE STYLE APPLICATION ---
        try:
            table.style = 'Table Grid'
        except KeyError:
            pass  # Style doesn't exist in template; skip styling to prevent crash
        # -----------------------------------
        
        table.autofit = False
        
        # Set column widths
        widths = [Cm(1.2), Cm(6.5), Cm(2.0), Cm(1.5), Cm(2.5), Cm(3.0)]
        for i, w in enumerate(widths): table.columns[i].width = w

        # Header Row
        for i, text in enumerate(headers):
            cell = table.rows[0].cells[i]
            cell.text = text
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if cell.paragraphs[0].runs:
                cell.paragraphs[0].runs[0].bold = True
            else:
                # If no runs exist, add one to make it bold
                run = cell.paragraphs[0].add_run(text)
                run.bold = True
                cell.text = "" # Clear duplicate text if run added manually
                cell.paragraphs[0].add_run(text).bold = True

        total_amt = 0
        for i, item in enumerate(cart_items):
            row = table.add_row().cells
            lp = item['List Price']
            disc = item['Discount']
            qty = item['Qty']
            net_rate = lp * (1 - disc/100)
            line_total = net_rate * qty
            total_amt += line_total
            
            desc = item['Description']
            if item.get('Make'): desc += f" ({item['Make']})"
            
            row[0].text = str(i+1)
            row[1].text = desc
            row[2].text = f"{qty:,.2f}"
            row[3].text = item['Display Unit'].split()[0]
            row[4].text = f"{net_rate:,.2f}"
            row[5].text = f"{line_total:,.2f}"
            
            for idx in [2, 4, 5]: row[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

        # Total Row
        row = table.add_row().cells
        row[1].text = "Total (Excl. GST)"
        row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if row[1].paragraphs[0].runs:
            row[1].paragraphs[0].runs[0].bold = True
        else:
             row[1].paragraphs[0].add_run("Total (Excl. GST)").bold = True
             row[1].text = ""
             row[1].paragraphs[0].add_run("Total (Excl. GST)").bold = True
        
        row[5].text = f"{total_amt:,.2f}"
        row[5].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if row[5].paragraphs[0].runs:
            row[5].paragraphs[0].runs[0].bold = True
        else:
            row[5].paragraphs[0].add_run(f"{total_amt:,.2f}").bold = True
            row[5].text = ""
            row[5].paragraphs[0].add_run(f"{total_amt:,.2f}").bold = True
        
        target_paragraph._p.addnext(table._tbl)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- DATA LOADER ---
@st.cache_data(show_spinner=True)
def load_data_from_files():
    all_dfs = []
    
    if not os.path.exists(DATA_DIR):
        return pd.DataFrame(), ["❌ Data folder missing"]

    excel_files = [os.path.join(DATA_DIR, f) for f in os.listdir(DATA_DIR) 
                   if f.lower().endswith(".xlsx") and not f.startswith("~$")]
    
    if not excel_files: return pd.DataFrame(), ["⚠️ No .xlsx files found in 'data' folder."]

    for file_path in excel_files:
        filename = os.path.basename(file_path)
        try:
            xls = pd.ExcelFile(file_path)
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet)
                
                # Simple Column Detection
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
                    clean_df['Display Unit'] = clean_df['UOM'] # Default
                    
                    clean_df = clean_df[clean_df['List Price'] > 0]
                    all_dfs.append(clean_df)
        except Exception as e:
            pass

    if not all_dfs: return pd.DataFrame(), ["⚠️ Could not read data from Excel files."]
    return pd.concat(all_dfs, ignore_index=True), []

# --- APP UI ---
catalog, logs = load_data_from_files()
if 'cart' not in st.session_state: st.session_state['cart'] = []

# SIDEBAR
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
        make = c2.text_input("Make")
        
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

# MAIN PAGE
st.title("📄 Quotation System")

if not st.session_state['cart']:
    st.info("Add items from the sidebar to start.")
    if logs: st.warning(logs[0])
else:
    # CART TABLE
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
        c1.write(f"{i+1}")
        c2.write(item['Description'])
        c3.write(item['Make'])
        c4.write(f"{item['Qty']:,.2f}")
        c5.write(item['Display Unit'])
        c6.write(f"₹ {tot:,.0f}")
        if c7.button("x", key=f"d{i}"):
            st.session_state['cart'].pop(i)
            st.rerun()

    st.divider()
    st.write(f"**Total (Excl. GST): ₹ {grand_tot:,.2f}**")
    st.markdown("---")

    # GENERATOR FORM
    st.subheader("2. Generate Official Quotation")
    
    selected_firm = st.selectbox("Select Template (Firm)", list(FIRM_MAPPING.keys()), key="firm_selector", on_change=update_defaults)
    
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
            client_data = {
                "client_name": client_name, "client_address": client_address,
                "ref_no": ref_no, "subject": subject
            }
            terms = {
                "price_term": p_term, "gst_term": g_term, "delivery_term": d_term,
                "freight_term": f_term, "payment_term": pay_term, "validity_term": val_term,
                "guarantee_term": guarantee
            }
            
            template_path = os.path.join(TEMPLATE_DIR, FIRM_MAPPING[selected_firm])
            docx_file = fill_template_docx(template_path, client_data, st.session_state['cart'], terms)
            
            st.download_button(
                label=f"Download Quote for {selected_firm}",
                data=docx_file,
                file_name=f"Quote_{selected_firm.replace(' ','_')}_{client_name[:5]}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )