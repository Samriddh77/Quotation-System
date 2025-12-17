import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime
from io import BytesIO

# Try importing python-docx (Handle error if not installed)
try:
    from docx import Document
    from docx.shared import Pt, Inches, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
except ImportError:
    st.error("❌ Library 'python-docx' is missing. Please run: pip install python-docx")
    st.stop()

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

# --- 2. WORD GENERATOR FUNCTION ---
def create_docx(client_data, cart_items, terms):
    doc = Document()
    
    # Styles
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    # --- HEADER SECTION ---
    # Ref and Date
    header_table = doc.add_table(rows=1, cols=2)
    header_table.autofit = True
    header_table.width = Inches(7.5)
    
    c1 = header_table.cell(0, 0)
    c1.text = f"Our Ref: {client_data['ref_no']}"
    
    c2 = header_table.cell(0, 1)
    c2.text = f"Date: {datetime.now().strftime('%d-%b-%Y')}"
    c2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    doc.add_paragraph() # Spacer

    # To Address
    p = doc.add_paragraph()
    p.add_run("To,\n").bold = True
    p.add_run(client_data['client_name'] + "\n")
    p.add_run(client_data['client_address'])
    
    doc.add_paragraph()

    # Subject
    p = doc.add_paragraph()
    runner = p.add_run(f"Sub: {client_data['subject']}")
    runner.bold = True
    runner.underline = True

    # Salutation
    doc.add_paragraph("Sirs,")
    doc.add_paragraph("We acknowledge with thanks the receipt of your above enquiry and are pleased to quote as under:-")

    # --- TERMS & CONDITIONS SECTION ---
    doc.add_paragraph().add_run("ANNEXURE I : PRICE SCHEDULE").bold = True
    
    p = doc.add_paragraph("Other Terms & Conditions are as under:")
    
    # Terms Table (Invisible borders for alignment) or List
    # Using list format as per your doc
    terms_list = [
        ("Price", terms['price_term']),
        ("GST", terms['gst_term']),
        ("Delivery", terms['delivery_term']),
        ("Freight", terms['freight_term']),
        ("Payment", terms['payment_term']),
        ("Validity", terms['validity_term']),
        ("Guarantee", terms['guarantee_term'])
    ]

    table_terms = doc.add_table(rows=len(terms_list), cols=2)
    table_terms.autofit = False
    table_terms.columns[0].width = Inches(1.5)
    table_terms.columns[1].width = Inches(5.0)
    
    for i, (k, v) in enumerate(terms_list):
        table_terms.cell(i, 0).text = f"{k} :"
        table_terms.cell(i, 1).text = v

    doc.add_paragraph()
    
    # --- CLOSING ---
    p = doc.add_paragraph()
    p.add_run("Thanking You\nYours Faithfully\n")
    p.add_run("For Electro World").bold = True
    
    doc.add_page_break()

    # --- ANNEXURE I (THE TABLE) ---
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h.add_run("ANNEXURE I: PRICE SCHEDULE").bold = True
    
    # Determine columns
    headers = ["S.No.", "Item Description", "Qty", "Unit", "Rate", "Amount", "Remark"]
    
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    table.autofit = False
    
    # Set Column Widths (Approximation)
    widths = [Cm(1.2), Cm(6.0), Cm(2.0), Cm(1.5), Cm(2.5), Cm(3.0), Cm(3.0)]
    for i, width in enumerate(widths):
        table.columns[i].width = width

    # Write Headers
    hdr_cells = table.rows[0].cells
    for i, h_text in enumerate(headers):
        hdr_cells[i].text = h_text
        hdr_cells[i].paragraphs[0].runs[0].bold = True
        hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Write Rows
    total_amt = 0
    for i, item in enumerate(cart_items):
        row_cells = table.add_row().cells
        
        # Calculations based on Terms (GST Inclusive vs Extra)
        # Assuming app data is always Base Price. 
        # Word doc "Rate" usually implies the Unit Net Rate.
        
        lp = item['List Price']
        disc = item['Discount']
        qty = item['Qty']
        
        net_rate = lp * (1 - disc/100)
        line_total = net_rate * qty
        total_amt += line_total
        
        desc = item['Description']
        if item.get('Make'): desc += f" ({item['Make']})"
        
        row_cells[0].text = str(i+1)
        row_cells[1].text = desc
        row_cells[2].text = f"{qty:,.2f}" if isinstance(qty, float) else str(qty)
        row_cells[3].text = item['Display Unit'].split()[0] # Just "Mtr" not "Mtr (5 Coils)"
        row_cells[4].text = f"{net_rate:,.2f}"
        row_cells[5].text = f"{line_total:,.2f}"
        row_cells[6].text = item.get('Remark', '')

    # Total Row
    row_cells = table.add_row().cells
    row_cells[1].text = "Total (Excl. GST)"
    row_cells[5].text = f"{total_amt:,.2f}"
    row_cells[1].paragraphs[0].runs[0].bold = True
    row_cells[5].paragraphs[0].runs[0].bold = True

    # Save to buffer
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- 3. DATA LOADER ---
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
                            ("ITEM" in r_str or "DESC" in r_str)):
                            header_idx = i
                            break
                    if header_idx == -1: continue

                    df = pd.read_excel(xls, sheet, skiprows=header_idx)
                    df.columns = [str(c).strip() for c in df.columns]
                    
                    name_col = next((c for c in df.columns if any(k in c.upper() for k in ["DESC", "PARTICULARS", "ITEM"])), df.columns[0])
                    
                    price_col = next((c for c in df.columns if "PER MTR" in c.upper()), None)
                    if not price_col:
                        price_col = next((c for c in df.columns if any(k in c.upper() for k in ["LP", "RATE", "PRICE"]) and "AMOUNT" not in c.upper()), None)

                    disc_col = next((c for c in df.columns if "DISC" in c.upper()), None)
                    coil_col = next((c for c in df.columns if "COIL" in c.upper() and ("LEN" in c.upper() or "MTR" in c.upper())), None)

                    if price_col:
                        clean_df = pd.DataFrame()
                        clean_df['Description'] = df[name_col].astype(str)
                        clean_df['List Price'] = df[price_col].apply(clean_price_value)
                        clean_df['Standard Discount'] = pd.to_numeric(df[disc_col], errors='coerce').fillna(0) if disc_col else 0
                        clean_df['Coil Length'] = df[coil_col].apply(clean_coil_len) if coil_col else 0.0
                        clean_df['Main Category'] = main_cat
                        clean_df['Sub Category'] = sheet
                        
                        if "UOM" in [c.upper() for c in df.columns]:
                            clean_df['UOM'] = df[next(c for c in df.columns if "UOM" in c.upper())].astype(str)
                        else:
                            clean_df['UOM'] = detect_uom(sheet, price_col)

                        clean_df = clean_df[clean_df['List Price'] > 0]
                        all_dfs.append(clean_df)
                        logs.append(f"✅ Loaded {main_cat} -> {sheet}")
                except: pass
        except: pass

    if not all_dfs: return pd.DataFrame(), logs
    return pd.concat(all_dfs, ignore_index=True), logs

# --- 4. UI START ---
catalog, logs = load_data_from_files()
if 'cart' not in st.session_state: st.session_state['cart'] = []

# SIDEBAR
with st.sidebar:
    st.title("🔧 Config")
    if st.button("🔄 Refresh Data"): st.cache_data.clear(); st.rerun()
    with st.expander("Logs"):
        for l in logs: st.write(l)

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
        std_disc = row['Standard Discount']
        uom = row['UOM']
        coil = row['Coil Length']
        
        st.info(f"Price: {std_price} / {uom}")
        
        calc_qty = 0
        disp_unit = uom
        if coil > 0 and "M" in uom.upper():
            mode = st.radio("Input", ["Coils", "Meters"], horizontal=True)
            if mode == "Coils":
                n = st.number_input("Coils", 1, 500)
                calc_qty = n * coil
                disp_unit = f"Mtr ({n} Coils)"
            else:
                n = st.number_input("Meters", 1, 10000)
                calc_qty = n
        else:
            calc_qty = st.number_input(f"Qty ({uom})", 1, 10000)

        c1, c2 = st.columns(2)
        disc = c1.number_input("Disc %", 0.0, 100.0, std_disc)
        make = c2.text_input("Make")
        remark = st.text_input("Remark")
        
        if st.button("Add"):
            st.session_state['cart'].append({
                'Main Category': sel_cat, 'Sub Category': sel_sub,
                'Description': sel_prod, 'Make': make, 'Remark': remark,
                'Qty': calc_qty, 'Display Unit': disp_unit,
                'List Price': std_price, 'Discount': disc, 'GST Rate': 18.0
            })
            st.success("Added")
            
    if st.button("Clear Cart"): st.session_state['cart'] = []; st.rerun()

# MAIN PAGE
st.title("📄 Quotation System")

if not st.session_state['cart']:
    st.info("Add items to start.")
else:
    # --- TABLE PREVIEW ---
    st.subheader("1. Item List")
    data = []
    grand_tot = 0
    for i, item in enumerate(st.session_state['cart']):
        net = item['List Price'] * (1 - item['Discount']/100)
        tot = (net * 1.18) * item['Qty']
        grand_tot += tot
        data.append({
            "No": i+1, "Desc": item['Description'], "Make": item['Make'],
            "Qty": item['Qty'], "Unit": item['Display Unit'],
            "Price": item['List Price'], "Disc": item['Discount'],
            "Total (Incl GST)": f"{tot:,.0f}"
        })
    st.dataframe(pd.DataFrame(data).set_index("No"))
    st.write(f"**Est. Grand Total: ₹ {grand_tot:,.2f}**")
    st.divider()

    # --- WORD DOC GENERATOR FORM ---
    st.subheader("2. Generate Official Quotation (.docx)")
    
    c1, c2 = st.columns(2)
    with c1:
        client_name = st.text_input("Client Name", placeholder="M/s Aneesh Commercial Pvt Ltd")
        ref_no = st.text_input("Reference No", value="CABLE/EW-001")
        subject = st.text_input("Subject", value="OFFER FOR CABLES / ELECTRICAL GOODS")
    with c2:
        client_address = st.text_area("Client Address", placeholder="Indore, MP")
    
    st.markdown("#### Terms & Conditions")
    tc1, tc2, tc3 = st.columns(3)
    p_term = tc1.text_input("Price Term", value="Nett")
    g_term = tc2.text_input("GST Term", value="Extra @ 18%")
    d_term = tc3.text_input("Delivery", value="Ex Stock / 7 Days")
    
    tc4, tc5, tc6 = st.columns(3)
    f_term = tc4.text_input("Freight", value="Ex Our Godown (Local Cartage Extra)")
    pay_term = tc5.text_input("Payment", value="100% Against Delivery")
    val_term = tc6.text_input("Validity", value="7 Days")
    
    guarantee = st.text_input("Guarantee", value="12 months from commissioning or 18 months from dispatch")

    if st.button("📥 Download Word Document"):
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
            
            docx_file = create_docx(client_data, st.session_state['cart'], terms)
            
            st.download_button(
                label="Click to Download .docx",
                data=docx_file,
                file_name=f"Quotation_{client_name[:10]}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )