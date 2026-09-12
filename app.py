import streamlit as st
import pypdf
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
import re
from collections import defaultdict
from datetime import datetime
import pandas as pd
import io

st.set_page_config(page_title="GST Universal Converter Pro", page_icon="🧾", layout="wide")

st.markdown("<h2 style='text-align: center; color: #1F4E79;'>🧾 Xenium Tyres - Universal GST Invoice to Excel & Google Sheets</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #555;'>सपोर्टेड: 1. Discount फॉर्मेट (Invoice 112) | 2. Multiprint फॉर्मेट | 3. Smart GST बिलिंग</p>", unsafe_allow_html=True)
st.write("---")

uploaded_pdf = st.file_uploader("📂 अपना इनवॉइस PDF यहाँ अपलोड करें", type=["pdf"])

def clean_num(s):
    try:
        val = re.sub(r'[^\d\.]', '', str(s))
        return float(val) if val and val != '.' else 0.0
    except:
        return 0.0

def sanitize_desc(desc):
    desc = re.sub(r'^(?:[\(\)\%0-9\.\s]|PCS|NOS|-)+\s*', '', desc, flags=re.IGNORECASE)
    desc = re.sub(r'[\(\)\%]+$', '', desc).strip()
    return desc

def parse_items_universal(text):
    items = []
    hsn_iter = list(re.finditer(r'\b(40\d{6})\b', text))
    if not hsn_iter:
        return items

    bottom_hsn_idx = text.find("HSN/SAC Taxable Value")
    if bottom_hsn_idx == -1:
        bottom_hsn_idx = text.find("HSN / SAC\nTaxable Value")
    if bottom_hsn_idx == -1:
        bottom_hsn_idx = text.find("HSN / SAC")
        if bottom_hsn_idx != -1 and bottom_hsn_idx < text.find("Sr."):
            bottom_hsn_idx = -1

    table_hsns = []
    for h in hsn_iter:
        if bottom_hsn_idx != -1 and h.start() >= bottom_hsn_idx:
            continue
        table_hsns.append(h)

    for idx, h_match in enumerate(table_hsns):
        hsn = int(h_match.group(1))
        h_start = h_match.start()
        h_end = h_match.end()

        prev_end = table_hsns[idx-1].end() if idx > 0 else 0
        desc_raw = text[prev_end:h_start]
        lines = [l.strip() for l in desc_raw.split('\n') if l.strip()]
        desc_parts = []
        for l in lines:
            if re.match(r'^\d+$', l) and len(l) <= 2: continue
            if any(k in l for k in ["Xenium", "Tyres", "Ahraura", "Mirzapur", "231301", "Paura", "Kaimur", "Bihar", "NEAR", "WARD", "Kuddi", "Sr.", "No.", "Name of Product", "HSN", "Qty", "Rate", "Taxable Value", "Disc", "Total", "% Amount"]):
                continue
            desc_parts.append(l)
        raw_desc = " ".join(desc_parts).strip()
        item_desc = sanitize_desc(raw_desc) or f"Tyre / Tube Item {idx+1}"

        next_start = table_hsns[idx+1].start() if idx + 1 < len(table_hsns) else len(text)
        post_raw = text[h_end:next_start]
        
        for sw in ["Total", "CGST", "SGST", "Round off", "HSN/SAC", "HSN / SAC"]:
            sw_pos = post_raw.find(sw)
            if sw_pos != -1:
                post_raw = post_raw[:sw_pos]

        tokens = post_raw.split()
        qty = 1.0
        uqc = "PCS-PIECES"
        rate = 0.0
        taxable = 0.0
        cgst_a = 0.0
        sgst_a = 0.0
        tot = 0.0

        unit_idx = -1
        for t_i, tok in enumerate(tokens):
            if "PCS" in tok.upper():
                uqc = "PCS-PIECES"
                unit_idx = t_i
                break
            elif "NOS" in tok.upper():
                uqc = "NOS-NUMBERS"
                unit_idx = t_i
                break

        nums = []
        for tok in tokens:
            val = re.sub(r'[^\d\.]', '', tok)
            if val and val != '.':
                try:
                    nums.append(float(val))
                except:
                    pass

        if len(nums) > 0:
            qty = nums[0]
            rem_nums = nums[1:]
        else:
            rem_nums = []

        if len(rem_nums) >= 7:
            rate = rem_nums[0]
            taxable = rem_nums[1]
            cgst_a = rem_nums[3]
            sgst_a = rem_nums[5]
            tot = rem_nums[6]
        elif len(rem_nums) >= 3:
            rate = rem_nums[0]
            taxable = rem_nums[2]
            cgst_a = round(taxable * 0.09, 2)
            sgst_a = round(taxable * 0.09, 2)
            tot = round(taxable + cgst_a + sgst_a, 2)
        elif len(rem_nums) == 2:
            rate = rem_nums[0]
            taxable = rem_nums[1]
            cgst_a = round(taxable * 0.09, 2)
            sgst_a = round(taxable * 0.09, 2)
            tot = round(taxable + cgst_a + sgst_a, 2)
        elif len(rem_nums) == 1:
            taxable = rem_nums[0]
            rate = taxable / qty if qty else taxable
            cgst_a = round(taxable * 0.09, 2)
            sgst_a = round(taxable * 0.09, 2)
            tot = round(taxable + cgst_a + sgst_a, 2)

        items.append({
            'desc': item_desc,
            'hsn': hsn,
            'qty': qty,
            'uqc': uqc,
            'rate': rate,
            'taxable': taxable,
            'cgst_r': 9.0,
            'cgst_a': cgst_a,
            'sgst_r': 9.0,
            'sgst_a': sgst_a,
            'total': tot
        })

    return items

def parse_pdf_document(file_bytes):
    reader = pypdf.PdfReader(io.BytesIO(file_bytes))
    invoices = []
    
    for page_idx, page in enumerate(reader.pages):
        text = page.extract_text()
        if not text or "TAX INVOICE" not in text:
            continue
            
        seller_gstin = ""
        sg_m = re.search(r'GSTIN\s*[:\|\s]?\s*\n?\s*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})', text[:500])
        if sg_m:
            seller_gstin = sg_m.group(1)

        inv_no_m = re.search(r'Invoice\s*No\.?\s*[:\|\s]?\s*(\d+)', text)
        if not inv_no_m:
            inv_no_m = re.search(r'Invoice\s*No\.?\s*\n\s*(\d+)', text)
        inv_no = int(inv_no_m.group(1)) if inv_no_m else (page_idx + 1)
        
        inv_date_m = re.search(r'Invoice\s*Date\s*[:\|\s]?\s*([0-9]{1,2}[\.\-\/][0-9A-Za-z]{2,3}[\.\-\/][0-9]{2,4})', text)
        if not inv_date_m:
            inv_date_m = re.search(r'Invoice\s*Date\s*\n\s*([0-9]{1,2}[\.\-\/][0-9A-Za-z]{2,3}[\.\-\/][0-9]{2,4})', text)
        if not inv_date_m:
            inv_date_m = re.search(r'(?:Date:?|ORIGINAL FOR RECIPIENT)\s*\n?\s*([0-9]{1,2}[\.\-\/][0-9A-Za-z]{2,3}[\.\-\/][0-9]{2,4})', text)
            
        inv_date = ""
        if inv_date_m:
            raw_d = inv_date_m.group(1).strip()
            try:
                if '.' in raw_d:
                    dt = datetime.strptime(raw_d, "%d.%m.%Y")
                    inv_date = dt.strftime("%d-%b-%Y")
                elif '-' in raw_d:
                    parts = raw_d.split('-')
                    if len(parts[1]) == 2:
                        dt = datetime.strptime(raw_d, "%d-%m-%Y")
                        inv_date = dt.strftime("%d-%b-%Y")
                    else:
                        inv_date = raw_d
                else:
                    inv_date = raw_d
            except:
                inv_date = raw_d

        cust_name = "Unregistered Consumer"
        cust_addr = ""
        cust_gstin = ""
        pos = "09-Uttar Pradesh"

        cust_blk = re.search(r'Customer Detail\s*\n(.*?)(?:Place of\s*Supply|TAX INVOICE|Invoice No|Due Date|Sr\.\s*\n?No)', text, re.DOTALL)
        if cust_blk:
            cb = cust_blk.group(1)
            gst_all = re.findall(r'[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}', cb)
            for g in gst_all:
                if g != seller_gstin:
                    cust_gstin = g
                    break
                    
            nm = re.search(r'(?:Name|M\/S|MS\/)\s*[:\|\s]?\s*\n?\s*([^\n]+)', cb, re.IGNORECASE)
            if nm:
                c_cand = nm.group(1).strip()
                if not any(k in c_cand.lower() for k in ["address", "phone", "gstin", "customer"]):
                    cust_name = c_cand
                    
            ad = re.search(r'Address\s*[:\|\s]?\s*\n?\s*([^\n]+(?:\n\s*[^\n]+)?)', cb, re.IGNORECASE)
            if ad:
                lines = [l.strip() for l in ad.group(1).split('\n') if l.strip() and not any(k in l.lower() for k in ['phone', 'gstin', 'place of'])]
                cust_addr = " ".join(lines)

        pos_m = re.search(r'Place of\s*Supply\s*[:\|\s]?\s*\n?\s*([^\n\(\)]+)\s*(?:\(\s*(\d+)\s*\))?', text)
        if pos_m:
            st_name = pos_m.group(1).strip()
            st_code = pos_m.group(2)
            if st_code:
                pos = f"{st_code.zfill(2)}-{st_name}"
            elif "bihar" in st_name.lower():
                pos = "10-Bihar"
            elif "uttar" in st_name.lower():
                pos = "09-Uttar Pradesh"
            else:
                pos = st_name

        taxable_m = re.search(r'Taxable Amount\s*[:\|\s]?\s*\n?\s*([\d,\.]+)', text)
        cgst_m = re.search(r'Add\s*:\s*CGST\s*[:\|\s]?\s*\n?\s*([\d,\.]+)', text)
        sgst_m = re.search(r'Add\s*:\s*SGST\s*[:\|\s]?\s*\n?\s*([\d,\.]+)', text)
        tot_m = re.search(r'Total Amount After Tax\s*[:\|\s]?\s*\n?\s*[₹\s]*([\d,\.]+)', text)
        if not tot_m:
            tot_m = re.search(r'Total\s*\n?\s*[\d\.]+\s*(?:PCS|NOS)\s*\n?\s*[\d,\.]+\s*\n?\s*[₹\s]*([\d,\.]+)', text)

        taxable_val = clean_num(taxable_m.group(1)) if taxable_m else 0.0
        cgst_val = clean_num(cgst_m.group(1)) if cgst_m else 0.0
        sgst_val = clean_num(sgst_m.group(1)) if sgst_m else 0.0
        tot_val = clean_num(tot_m.group(1)) if tot_m else 0.0

        items = parse_items_universal(text)
        if taxable_val == 0.0 and items:
            taxable_val = round(sum(it['taxable'] for it in items), 2)
        if tot_val == 0.0 and items:
            tot_val = round(sum(it['total'] for it in items), 2)

        invoices.append({
            'invoice_no': inv_no,
            'invoice_date': inv_date,
            'customer_name': cust_name,
            'address': cust_addr,
            'cust_gstin': cust_gstin,
            'pos': pos,
            'taxable_val': taxable_val,
            'cgst_val': cgst_val,
            'sgst_val': sgst_val,
            'tot_val': tot_val,
            'items': items
        })
        
    return invoices

def generate_full_gstr1_excel(invoices):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    b2b_invs = [x for x in invoices if x['cust_gstin']]
    b2c_invs = [x for x in invoices if not x['cust_gstin']]
    
    # 1. b2b,sez,de
    ws_b2b = wb.create_sheet('b2b,sez,de')
    ws_b2b.append(['Summary For B2B(4)'] + [None]*11 + ['HELP'])
    ws_b2b.append(['No. of Recipients', None, 'No. of Invoices', None, 'Total Invoice Value', None, None, None, None, None, None, 'Total Taxable Value', 'Total Cess'])
    ws_b2b.append([
        len(set(x['cust_gstin'] for x in b2b_invs)), None,
        len(b2b_invs), None,
        round(sum(x['tot_val'] for x in b2b_invs), 2), None, None, None, None, None, None,
        round(sum(x['taxable_val'] for x in b2b_invs), 2), 0.0
    ])
    ws_b2b.append(['GSTIN/UIN of Recipient', 'Receiver Name', 'Invoice Number', 'Invoice date', 'Invoice Value', 'Place Of Supply', 'Reverse Charge', 'Applicable % of Tax Rate', 'Invoice Type', 'E-Commerce GSTIN', 'Rate', 'Taxable Value', 'Cess Amount'])
    for inv in b2b_invs:
        ws_b2b.append([inv['cust_gstin'], inv['customer_name'], inv['invoice_no'], inv['invoice_date'], inv['tot_val'], inv['pos'], 'N', None, 'Regular B2B', None, 18.0, inv['taxable_val'], 0.0])

    # 2. b2cl
    ws_b2cl = wb.create_sheet('b2cl')
    ws_b2cl.append(['Summary For B2CL(5)'] + [None]*7 + ['HELP'])
    ws_b2cl.append(['No. of Invoices', None, 'Total Inv Value', None, None, None, 'Total Taxable Value', 'Total Cess', None])
    ws_b2cl.append([0, None, 0.0, None, None, None, 0.0, 0.0, None])
    ws_b2cl.append(['Invoice Number', 'Invoice date', 'Invoice Value', 'Place Of Supply', 'Applicable % of Tax Rate', 'Rate', 'Taxable Value', 'Cess Amount', 'E-Commerce GSTIN'])

    # 3. b2cs
    ws_b2cs = wb.create_sheet('b2cs')
    ws_b2cs.append(['Summary For B2CS(7)'] + [None]*5 + ['HELP'])
    ws_b2cs.append([None, None, None, None, 'Total Taxable  Value', 'Total Cess', None])
    tot_b2c_tax = round(sum(x['taxable_val'] for x in b2c_invs), 2)
    ws_b2cs.append([None, None, None, None, tot_b2c_tax, 0.0, None])
    ws_b2cs.append(['Type', 'Place Of Supply', 'Applicable % of Tax Rate', 'Rate', 'Taxable Value', 'Cess Amount', 'E-Commerce GSTIN'])
    if tot_b2c_tax > 0:
        pos_b2c = b2c_invs[0]['pos'] if b2c_invs else '09-Uttar Pradesh'
        ws_b2cs.append(['OE', pos_b2c, None, 18.0, tot_b2c_tax, 0.0, None])

    # 4 & 5. HSN Sheets
    def add_hsn_sheet(sheet_name, inv_list):
        ws = wb.create_sheet(sheet_name)
        ws.append(['Summary For HSN(12)'] + [None]*8 + ['HELP', None])
        ws.append(['No. of HSN', None, None, None, 'Total Value', None, 'Total Taxable Value', 'Total Integrated Tax', 'Total Central Tax', 'Total State/UT Tax', 'Total Cess'])
        
        group = defaultdict(lambda: {'desc': set(), 'qty': 0.0, 'total': 0.0, 'taxable': 0.0, 'cgst': 0.0, 'sgst': 0.0})
        for inv in inv_list:
            for it in inv['items']:
                key = (it['hsn'], it['uqc'])
                group[key]['desc'].add(it['desc'])
                group[key]['qty'] += it['qty']
                group[key]['total'] += it['total']
                group[key]['taxable'] += it['taxable']
                group[key]['cgst'] += it['cgst_a']
                group[key]['sgst'] += it['sgst_a']
                
        ws.append([
            len(group), None, None, None,
            round(sum(v['total'] for v in group.values()), 2), None,
            round(sum(v['taxable'] for v in group.values()), 2), 0.0,
            round(sum(v['cgst'] for v in group.values()), 2),
            round(sum(v['sgst'] for v in group.values()), 2), 0.0
        ])
        ws.append(['HSN', 'Description', 'UQC', 'Total Quantity', 'Total Value', 'Rate', 'Taxable Value', 'Integrated Tax Amount', 'Central Tax Amount', 'State/UT Tax Amount', 'Cess Amount'])
        for (hsn_code, uqc_unit), data in sorted(group.items()):
            desc_str = ", ".join(sorted(data['desc']))
            ws.append([int(hsn_code), desc_str, uqc_unit, round(data['qty'], 2), round(data['total'], 2), 18.0, round(data['taxable'], 2), 0.0, round(data['cgst'], 2), round(data['sgst'], 2), 0.0])

    add_hsn_sheet('hsn(b2b)', b2b_invs)
    add_hsn_sheet('hsn(b2c)', b2c_invs)

    # 6. docs
    ws_docs = wb.create_sheet('docs')
    ws_docs.append(['Summary of documents issued during the tax period (13)'] + [None]*3 + ['HELP'])
    ws_docs.append([None, None, None, 'Total Number', 'Total Cancelled'])
    inv_nums = [x['invoice_no'] for x in invoices if x['invoice_no']]
    ws_docs.append([None, None, None, len(inv_nums), 0])
    ws_docs.append(['Nature of Document', 'Sr. No. From', 'Sr. No. To', 'Total Number', 'Cancelled'])
    ws_docs.append(['Invoices for outward supply', min(inv_nums) if inv_nums else 1, max(inv_nums) if inv_nums else len(inv_nums), len(inv_nums), 0])

    # 7. Invoice_Register
    ws_reg = wb.create_sheet('Invoice_Register')
    headers = ['Invoice No', 'Date', 'Customer Name', 'Address', 'GSTIN', 'Place Of Supply', 'Item Description', 'HSN', 'Qty', 'Unit', 'Rate', 'Taxable Value', 'CGST', 'SGST', 'Total']
    ws_reg.append(headers)
    for c in range(1, len(headers)+1):
        cell = ws_reg.cell(1, c)
        cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        cell.font = Font(name="Calibri", bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center")

    for inv in invoices:
        for it in inv['items']:
            ws_reg.append([
                inv['invoice_no'], inv['invoice_date'], inv['customer_name'],
                inv['address'], inv['cust_gstin'] or 'Unregistered', inv['pos'],
                it['desc'], it['hsn'], it['qty'], it['uqc'], it['rate'],
                it['taxable'], it['cgst_a'], it['sgst_a'], it['total']
            ])

    return wb

if uploaded_pdf and st.button("🚀 Convert to GSTR-1 Excel & Google Sheets", type="primary", use_container_width=True):
    with st.spinner("PDF से बिल डेटा स्कैन और प्रोसेस किया जा रहा है..."):
        invoices = parse_pdf_document(uploaded_pdf.read())
        
        if not invoices:
            st.error("कोई बिल नहीं पढ़ा जा सका। कृपया सही इनवॉइस PDF चुनें।")
        else:
            wb = generate_full_gstr1_excel(invoices)
            
            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_buffer.seek(0)
            
            flat_rows = []
            for inv in invoices:
                for it in inv['items']:
                    flat_rows.append({
                        'Invoice No': inv['invoice_no'],
                        'Date': inv['invoice_date'],
                        'Customer Name': inv['customer_name'],
                        'Customer GSTIN': inv['cust_gstin'] or 'Unregistered',
                        'Place Of Supply': inv['pos'],
                        'Item Description': it['desc'],
                        'HSN': it['hsn'],
                        'Quantity': it['qty'],
                        'Unit': it['uqc'],
                        'Rate': it['rate'],
                        'Taxable Value': it['taxable'],
                        'CGST (9%)': it['cgst_a'],
                        'SGST (9%)': it['sgst_a'],
                        'Total Amount': it['total']
                    })
            df = pd.DataFrame(flat_rows)
            csv_data = df.to_csv(index=False).encode('utf-8')

            b2b_cnt = sum(1 for x in invoices if x['cust_gstin'])
            b2c_cnt = sum(1 for x in invoices if not x['cust_gstin'])
            total_sales = sum(x['tot_val'] for x in invoices)

            st.success(f"✅ सफलता! कुल {len(invoices)} बिल एक्सट्रैक्ट हुए (B2B: {b2b_cnt} | B2C: {b2c_cnt}) | कुल बिक्री: ₹{total_sales:,.2f}")
            
            st.subheader("📥 1. फ़ाइल डाउनलोड विकल्प")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.download_button(
                    label="📊 Download GSTR-1 Excel (.xlsx)",
                    data=excel_buffer,
                    file_name="GSTR1_Filled_Ready.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
            with col2:
                st.download_button(
                    label="📄 Download for Google Sheets (.csv)",
                    data=csv_data,
                    file_name="Invoices_For_GoogleSheets.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            with col3:
                st.link_button(
                    label="🌐 Open Google Sheets",
                    url="https://sheets.new",
                    use_container_width=True
                )

            st.markdown("---")
            st.subheader("👁️ 2. लाइव Google Sheets व्यू (सीधा स्क्रीन पर)")
            
            tab1, tab2, tab3 = st.tabs(["📋 All Invoices Register", "🏢 B2B Invoices", "📊 B2CS & HSN Summary"])
            with tab1:
                st.dataframe(df, use_container_width=True)
            with tab2:
                b2b_df = df[df['Customer GSTIN'] != 'Unregistered']
                if not b2b_df.empty:
                    st.dataframe(b2b_df, use_container_width=True)
                else:
                    st.info("इस PDF में कोई B2B ग्राहक (GSTIN वाला) नहीं मिला।")
            with tab3:
                hsn_summary = df.groupby(['HSN', 'Unit']).agg({'Quantity': 'sum', 'Taxable Value': 'sum', 'Total Amount': 'sum'}).reset_index()
                st.dataframe(hsn_summary, use_container_width=True)

            st.info("💡 **Google Sheets में 1-क्लिक से कैसे खोलें:**\n1. ऊपर **'Download for Google Sheets (.csv)'** या Excel पर क्लिक करें।\n2. **'Open Google Sheets'** पर टैप करें।\n3. Google Sheets में **File ➔ Import ➔ Upload** से फ़ाइल चुनें; पूरा डेटा तुरंत शीट में आ जाएगा।")
