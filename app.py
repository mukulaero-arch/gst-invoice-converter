import streamlit as st
import pypdf
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
import re
from collections import defaultdict
from datetime import datetime
import io

st.set_page_config(page_title="GST Invoice Converter Pro", page_icon="🧾", layout="wide")

st.markdown("<h2 style='text-align: center; color: #1F4E79;'>🧾 Xenium Tyres - Universal GST PDF to Excel / Sheets</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: gray;'>सपोर्टेड: Multiprint PDF एवं Smart GST Billing PDF</p>", unsafe_allow_html=True)
st.write("---")

uploaded_pdf = st.file_uploader("📂 1. इनवॉइस PDF अपलोड करें (कोई भी फॉर्मेट)", type=["pdf"])
uploaded_tpl = st.file_uploader("📑 2. GSTR-1 Template Excel अपलोड करें (GSTR1.xlsx)", type=["xlsx"])

def parse_any_pdf(file_bytes):
    reader = pypdf.PdfReader(io.BytesIO(file_bytes))
    invoices = []
    
    for page_idx, page in enumerate(reader.pages):
        text = page.extract_text()
        if not text or "TAX INVOICE" not in text:
            continue
            
        # 1. Invoice No
        inv_no_m = re.search(r'Invoice\s*No\.?\s*[:\|\s]?\s*(\d+)', text)
        if not inv_no_m:
            inv_no_m = re.search(r'Invoice\s*No\.?\s*\n\s*(\d+)', text)
        inv_no = int(inv_no_m.group(1)) if inv_no_m else (page_idx + 1)
        
        # 2. Invoice Date
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
                
        # 3. Customer Details & GSTIN
        cust_name = "Unregistered Consumer"
        cust_addr = ""
        cust_gstin = ""
        
        cust_blk = re.search(r'Customer Detail\s*\n(.*?)(?:Place of\s*Supply|TAX INVOICE|Invoice No|Due Date)', text, re.DOTALL)
        if cust_blk:
            cb = cust_blk.group(1)
            gst_all = re.findall(r'[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}', cb)
            for g in gst_all:
                if g != "09KQTPS5380E1ZW":
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
                
        # 4. Totals
        taxable_m = re.search(r'Taxable Amount\s*[:\|\s]?\s*\n?\s*([\d,\.]+)', text)
        cgst_m = re.search(r'Add\s*:\s*CGST\s*[:\|\s]?\s*\n?\s*([\d,\.]+)', text)
        sgst_m = re.search(r'Add\s*:\s*SGST\s*[:\|\s]?\s*\n?\s*([\d,\.]+)', text)
        tot_m = re.search(r'Total Amount After Tax\s*[:\|\s]?\s*\n?\s*[₹\s]*([\d,\.]+)', text)
        
        taxable_val = float(taxable_m.group(1).replace(',', '')) if taxable_m else 0.0
        cgst_val = float(cgst_m.group(1).replace(',', '')) if cgst_m else 0.0
        sgst_val = float(sgst_m.group(1).replace(',', '')) if sgst_m else 0.0
        tot_val = float(tot_m.group(1).replace(',', '')) if tot_m else (taxable_val + cgst_val + sgst_val)
        
        # 5. Items Extraction
        items_area_m = re.search(r'Sr\.\s*\n?No\..*?(?:Total\s*\n?\s*[\d\.]+\s*(?:PCS|NOS))', text, re.DOTALL)
        items = []
        if items_area_m:
            area = items_area_m.group(0)
            hsn_matches = list(re.finditer(r'(40\d{6})\s*\n?\s*(\d+(?:\.\d+)?)\s*(PCS|NOS)?\s*\n?\s*([\d,\.]+)\s*\n?\s*([\d,\.]+)\s*\n?\s*([\d,\.]+)\s*\n?\s*([\d,\.]+)\s*\n?\s*([\d,\.]+)\s*\n?\s*([\d,\.]+)\s*\n?\s*([\d,\.]+)', area))
            for idx_h, hm in enumerate(hsn_matches):
                hsn, qty, uqc_unit, rate, tax_amt, cgst_r, cgst_a, sgst_r, sgst_a, item_tot = hm.groups()
                
                h_start = hm.start()
                start_search = hsn_matches[idx_h - 1].end() if idx_h > 0 else 0
                desc_text = area[start_search:h_start]
                d_lines = [l.strip() for l in desc_text.split('\n') if l.strip()]
                clean_words = []
                for l in d_lines:
                    if l.isdigit() and len(l) <= 2: continue
                    if any(bad in l for bad in ["Xenium", "Tyres", "Ahraura", "Mirzapur", "231301", "Sr.", "No.", "Name of Product", "HSN", "Qty", "Rate", "Taxable Value", "CGST", "SGST", "Total", "% Amount"]): continue
                    clean_words.append(l)
                item_desc = " ".join(clean_words).strip() or f"Tyre / Tube (HSN {hsn})"
                uqc = "NOS-NUMBERS" if (uqc_unit and "NOS" in uqc_unit) else "PCS-PIECES"
                
                items.append({
                    'desc': item_desc,
                    'hsn': int(hsn),
                    'qty': float(qty),
                    'uqc': uqc,
                    'rate': float(rate.replace(',', '')),
                    'taxable': float(tax_amt.replace(',', '')),
                    'cgst_r': float(cgst_r.replace(',', '')),
                    'cgst_a': float(cgst_a.replace(',', '')),
                    'sgst_r': float(sgst_r.replace(',', '')),
                    'sgst_a': float(sgst_a.replace(',', '')),
                    'total': float(item_tot.replace(',', ''))
                })
                
        invoices.append({
            'invoice_no': inv_no,
            'invoice_date': inv_date,
            'customer_name': cust_name,
            'address': cust_addr,
            'cust_gstin': cust_gstin,
            'pos': "09-Uttar Pradesh",
            'taxable_val': taxable_val,
            'cgst_val': cgst_val,
            'sgst_val': sgst_val,
            'tot_val': tot_val,
            'items': items
        })
        
    return invoices

if uploaded_pdf and st.button("🚀 Convert to GSTR-1 Excel", type="primary", use_container_width=True):
    with st.spinner("दोनों PDF फॉर्मेट्स को प्रोसेस किया जा रहा है..."):
        pdf_bytes = uploaded_pdf.read()
        invoices = parse_any_pdf(pdf_bytes)
        
        if not invoices:
            st.error("कोई बिल नहीं मिला। कृपया सही PDF फ़ाइल चुनें।")
        else:
            if uploaded_tpl:
                wb = openpyxl.load_workbook(uploaded_tpl)
            else:
                st.error("कृपया अपना GSTR1 (5).xlsx टेम्पलेट अपलोड करें।")
                st.stop()
                
            b2b_invs = [x for x in invoices if x['cust_gstin']]
            b2c_invs = [x for x in invoices if not x['cust_gstin']]
            
            # 1. B2B Sheet
            if 'b2b,sez,de' in wb.sheetnames:
                ws_b2b = wb['b2b,sez,de']
                while ws_b2b.max_row > 4: ws_b2b.delete_rows(5)
                
                unique_rec = len(set(x['cust_gstin'] for x in b2b_invs))
                tot_b2b_val = round(sum(x['tot_val'] for x in b2b_invs), 2)
                tot_b2b_tax = round(sum(x['taxable_val'] for x in b2b_invs), 2)
                
                ws_b2b.cell(3, 1, unique_rec)
                ws_b2b.cell(3, 3, len(b2b_invs))
                ws_b2b.cell(3, 5, tot_b2b_val)
                ws_b2b.cell(3, 12, tot_b2b_tax)
                ws_b2b.cell(3, 13, 0.0)
                
                r = 5
                for inv in b2b_invs:
                    ws_b2b.cell(r, 1, inv['cust_gstin'])
                    ws_b2b.cell(r, 2, inv['customer_name'])
                    ws_b2b.cell(r, 3, inv['invoice_no'])
                    ws_b2b.cell(r, 4, inv['invoice_date'])
                    ws_b2b.cell(r, 5, inv['tot_val'])
                    ws_b2b.cell(r, 6, inv['pos'])
                    ws_b2b.cell(r, 7, 'N')
                    ws_b2b.cell(r, 9, 'Regular B2B')
                    ws_b2b.cell(r, 11, 18.0)
                    ws_b2b.cell(r, 12, inv['taxable_val'])
                    ws_b2b.cell(r, 13, 0.0)
                    r += 1

            # 2. B2CS Sheet
            if 'b2cs' in wb.sheetnames:
                ws_b2cs = wb['b2cs']
                while ws_b2cs.max_row > 4: ws_b2cs.delete_rows(5)
                
                tot_b2c_tax = round(sum(x['taxable_val'] for x in b2c_invs), 2)
                ws_b2cs.cell(3, 5, tot_b2c_tax)
                ws_b2cs.cell(3, 6, 0.0)
                if tot_b2c_tax > 0:
                    ws_b2cs.cell(5, 1, 'OE')
                    ws_b2cs.cell(5, 2, '09-Uttar Pradesh')
                    ws_b2cs.cell(5, 4, 18.0)
                    ws_b2cs.cell(5, 5, tot_b2c_tax)
                    ws_b2cs.cell(5, 6, 0.0)

            # 3. HSN (B2B)
            hsn_b2b = defaultdict(lambda: {'desc': set(), 'uqc': 'PCS-PIECES', 'qty': 0.0, 'total': 0.0, 'taxable': 0.0, 'cgst': 0.0, 'sgst': 0.0})
            for inv in b2b_invs:
                for it in inv['items']:
                    h = it['hsn']
                    hsn_b2b[h]['desc'].add(it['desc'])
                    hsn_b2b[h]['qty'] += it['qty']
                    hsn_b2b[h]['uqc'] = it['uqc']
                    hsn_b2b[h]['total'] += it['total']
                    hsn_b2b[h]['taxable'] += it['taxable']
                    hsn_b2b[h]['cgst'] += it['cgst_a']
                    hsn_b2b[h]['sgst'] += it['sgst_a']
                    
            if 'hsn(b2b)' in wb.sheetnames:
                ws_hb = wb['hsn(b2b)']
                while ws_hb.max_row > 4: ws_hb.delete_rows(5)
                ws_hb.cell(3, 1, len(hsn_b2b))
                ws_hb.cell(3, 5, round(sum(v['total'] for v in hsn_b2b.values()), 2))
                ws_hb.cell(3, 7, round(sum(v['taxable'] for v in hsn_b2b.values()), 2))
                ws_hb.cell(3, 8, 0.0)
                ws_hb.cell(3, 9, round(sum(v['cgst'] for v in hsn_b2b.values()), 2))
                ws_hb.cell(3, 10, round(sum(v['sgst'] for v in hsn_b2b.values()), 2))
                ws_hb.cell(3, 11, 0.0)
                
                r = 5
                for hsn, data in sorted(hsn_b2b.items()):
                    ws_hb.cell(r, 1, int(hsn))
                    ws_hb.cell(r, 2, ", ".join(sorted(data['desc'])))
                    ws_hb.cell(r, 3, data['uqc'])
                    ws_hb.cell(r, 4, data['qty'])
                    ws_hb.cell(r, 5, round(data['total'], 2))
                    ws_hb.cell(r, 6, 18.0)
                    ws_hb.cell(r, 7, round(data['taxable'], 2))
                    ws_hb.cell(r, 8, 0.0)
                    ws_hb.cell(r, 9, round(data['cgst'], 2))
                    ws_hb.cell(r, 10, round(data['sgst'], 2))
                    ws_hb.cell(r, 11, 0.0)
                    r += 1

            # 4. HSN (B2C)
            hsn_b2c = defaultdict(lambda: {'desc': set(), 'uqc': 'PCS-PIECES', 'qty': 0.0, 'total': 0.0, 'taxable': 0.0, 'cgst': 0.0, 'sgst': 0.0})
            for inv in b2c_invs:
                for it in inv['items']:
                    h = it['hsn']
                    hsn_b2c[h]['desc'].add(it['desc'])
                    hsn_b2c[h]['qty'] += it['qty']
                    hsn_b2c[h]['uqc'] = it['uqc']
                    hsn_b2c[h]['total'] += it['total']
                    hsn_b2c[h]['taxable'] += it['taxable']
                    hsn_b2c[h]['cgst'] += it['cgst_a']
                    hsn_b2c[h]['sgst'] += it['sgst_a']
                    
            if 'hsn(b2c)' in wb.sheetnames:
                ws_hc = wb['hsn(b2c)']
                while ws_hc.max_row > 4: ws_hc.delete_rows(5)
                ws_hc.cell(3, 1, len(hsn_b2c))
                ws_hc.cell(3, 5, round(sum(v['total'] for v in hsn_b2c.values()), 2))
                ws_hc.cell(3, 7, round(sum(v['taxable'] for v in hsn_b2c.values()), 2))
                ws_hc.cell(3, 8, 0.0)
                ws_hc.cell(3, 9, round(sum(v['cgst'] for v in hsn_b2c.values()), 2))
                ws_hc.cell(3, 10, round(sum(v['sgst'] for v in hsn_b2c.values()), 2))
                ws_hc.cell(3, 11, 0.0)
                
                r = 5
                for hsn, data in sorted(hsn_b2c.items()):
                    ws_hc.cell(r, 1, int(hsn))
                    ws_hc.cell(r, 2, ", ".join(sorted(data['desc'])))
                    ws_hc.cell(r, 3, data['uqc'])
                    ws_hc.cell(r, 4, data['qty'])
                    ws_hc.cell(r, 5, round(data['total'], 2))
                    ws_hc.cell(r, 6, 18.0)
                    ws_hc.cell(r, 7, round(data['taxable'], 2))
                    ws_hc.cell(r, 8, 0.0)
                    ws_hc.cell(r, 9, round(data['cgst'], 2))
                    ws_hc.cell(r, 10, round(data['sgst'], 2))
                    ws_hc.cell(r, 11, 0.0)
                    r += 1

            # 5. Docs Sheet
            if 'docs' in wb.sheetnames:
                ws_docs = wb['docs']
                inv_nums = [x['invoice_no'] for x in invoices if x['invoice_no']]
                ws_docs.cell(3, 4, len(inv_nums))
                ws_docs.cell(3, 5, 0)
                ws_docs.cell(5, 1, 'Invoices for outward supply')
                ws_docs.cell(5, 2, min(inv_nums) if inv_nums else 1)
                ws_docs.cell(5, 3, max(inv_nums) if inv_nums else len(inv_nums))
                ws_docs.cell(5, 4, len(inv_nums))
                ws_docs.cell(5, 5, 0)

            # 6. Detailed Invoice Register Sheet
            if "Invoice_Register" in wb.sheetnames:
                del wb["Invoice_Register"]
            ws_reg = wb.create_sheet(title="Invoice_Register")
            headers = ['Inv No', 'Date', 'Customer Name', 'Address', 'GSTIN', 'Item Desc', 'HSN', 'Qty', 'Unit', 'Rate', 'Taxable', 'CGST', 'SGST', 'Total']
            ws_reg.append(headers)
            for c in range(1, len(headers)+1):
                ws_reg.cell(1, c).fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
                ws_reg.cell(1, c).font = Font(bold=True, color="FFFFFF")
                ws_reg.cell(1, c).alignment = Alignment(horizontal="center")
                
            for inv in invoices:
                for it in inv['items']:
                    ws_reg.append([
                        inv['invoice_no'], inv['invoice_date'], inv['customer_name'],
                        inv['address'], inv['cust_gstin'] or 'Unregistered', it['desc'],
                        it['hsn'], it['qty'], it['uqc'], it['rate'], it['taxable'],
                        it['cgst_a'], it['sgst_a'], it['total']
                    ])

            out_buf = io.BytesIO()
            wb.save(out_buf)
            out_buf.seek(0)
            
            st.success(f"✅ कुल {len(invoices)} बिल सफलतापूर्वक भरे गए! (B2B: {len(b2b_invs)}, B2C: {len(b2c_invs)})")
            
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📥 Download Filled GSTR-1 Excel (.xlsx)",
                    data=out_buf,
                    file_name="GSTR1_Filled_Ready.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
            with col2:
                st.link_button(
                    label="📂 Open in Google Sheets",
                    url="https://sheets.new",
                    use_container_width=True
                )

            st.info("💡 **Google Sheets में खोलने का तरीक़ा:** ऊपर 'Download' करके फ़ाइल सेव करें, फिर 'Open in Google Sheets' पर क्लिक करें और **File ➔ Import ➔ Upload** से इस फ़ाइल को चुन लें।")
