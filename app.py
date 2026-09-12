import streamlit as st
import pypdf
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
import re
from collections import defaultdict
import io

st.set_page_config(page_title="GST Smart Converter", page_icon="📊", layout="wide")

st.markdown("<h2 style='text-align: center; color: #1F4E79;'>📊 Xenium Tyres - Smart GST Invoice to Excel / Sheets</h2>", unsafe_allow_html=True)
st.write("---")

uploaded_pdf = st.file_uploader("📂 1. अपने PDF बिल अपलोड करें", type=["pdf"])
uploaded_tpl = st.file_uploader("📑 2. GSTR-1 Excel Template अपलोड करें (वैकल्पिक)", type=["xlsx"])

if uploaded_pdf and st.button("🚀 Convert to GSTR-1 Excel", type="primary", use_container_width=True):
    with st.spinner("बिल प्रोसेस किए जा रहे हैं..."):
        reader = pypdf.PdfReader(uploaded_pdf)
        invoices = []

        for page in reader.pages:
            text = page.extract_text()
            if not text or "TAX INVOICE" not in text:
                continue

            inv_no_m = re.search(r'Invoice No\.\s*(\d+)', text)
            inv_date_m = re.search(r'Invoice Date\s*([\d\.]+)', text)
            cust_block = re.search(r'Customer Detail\s*\n(.*?)(?:TAX INVOICE|Invoice No)', text, re.DOTALL)
            
            c_name, c_addr, c_gstin = "Unregistered Consumer", "", ""
            if cust_block:
                cb = cust_block.group(1)
                nm = re.search(r'Name:\s*([^\n]+)', cb)
                ad = re.search(r'Address:\s*([^\n]+)', cb)
                gs = re.search(r'GSTIN:\s*([^\n]*)', cb)
                if nm: c_name = nm.group(1).strip()
                if ad: c_addr = ad.group(1).strip()
                if gs:
                    val = gs.group(1).strip()
                    if val and val != '-' and len(val) >= 15:
                        c_gstin = val

            pos_m = re.search(r'Place of Supply\s*([^\n]+)', text)
            pos = pos_m.group(1).strip() if pos_m else "09-Uttar Pradesh"

            taxable_amt_m = re.search(r'Taxable Amount\s*([\d,\.]+)', text)
            cgst_m = re.search(r'Add:\s*CGST\s*([\d,\.]+)', text)
            sgst_m = re.search(r'Add:\s*SGST\s*([\d,\.]+)', text)
            total_m = re.search(r'Total Amount After Tax\s*[₹\s]*([\d,\.]+)', text)

            taxable_val = float(taxable_amt_m.group(1).replace(',', '')) if taxable_amt_m else 0.0
            cgst_val = float(cgst_m.group(1).replace(',', '')) if cgst_m else 0.0
            sgst_val = float(sgst_m.group(1).replace(',', '')) if sgst_m else 0.0
            tot_val = float(total_m.group(1).replace(',', '')) if total_m else 0.0

            # आइटम निकालना
            items_part = re.search(r'% Amount % Amount\s*\n(.*?)\nTotal\s+\d+\s+PCS', text, re.DOTALL)
            items = []
            if items_part:
                item_blocks = re.split(r'\n(?=\d+\s+)', items_part.group(1).strip())
                for ib in item_blocks:
                    m = re.search(r'(\d+)\s+(.+?)\s+(\d{6,8})\s+(\d+(?:\.\d+)?)\s+([A-Z]+)\s+([\d,\.]+)\s+([\d,\.]+)\s+([\d,\.]+)\s+([\d,\.]+)\s+([\d,\.]+)\s+([\d,\.]+)\s+([\d,\.]+)', ib, re.DOTALL)
                    if m:
                        items.append({
                            'desc': ' '.join(m.group(2).split()),
                            'hsn': m.group(3),
                            'qty': float(m.group(4)),
                            'uqc': m.group(5) + '-PIECES' if m.group(5) == 'PCS' else m.group(5),
                            'rate': float(m.group(6).replace(',', '')),
                            'taxable_value': float(m.group(7).replace(',', '')),
                            'cgst_amount': float(m.group(9).replace(',', '')),
                            'sgst_amount': float(m.group(11).replace(',', '')),
                            'total': float(m.group(12).replace(',', ''))
                        })

            invoices.append({
                'invoice_no': int(inv_no_m.group(1)) if inv_no_m else 0,
                'invoice_date': inv_date_m.group(1) if inv_date_m else "",
                'customer_name': c_name,
                'address': c_addr,
                'cust_gstin': c_gstin,
                'pos': pos,
                'taxable_val': taxable_val,
                'cgst': cgst_val,
                'sgst': sgst_val,
                'total': tot_val,
                'items': items
            })

        # वर्कबुक लोड या क्रिएट करें
        if uploaded_tpl:
            wb = openpyxl.load_workbook(uploaded_tpl)
        else:
            wb = openpyxl.Workbook()
            wb.active.title = "b2cs"

        # B2B vs B2C अलग करें
        b2b_invs = [x for x in invoices if x['cust_gstin']]
        b2c_invs = [x for x in invoices if not x['cust_gstin']]

        # 1. B2CS शीट अपडेट
        if 'b2cs' in wb.sheetnames:
            ws_b2cs = wb['b2cs']
            while ws_b2cs.max_row > 4: ws_b2cs.delete_rows(5)
            t_taxable = round(sum(x['taxable_val'] for x in b2c_invs), 2)
            ws_b2cs.cell(3, 5, t_taxable)
            ws_b2cs.cell(3, 6, 0.0)
            if t_taxable > 0:
                ws_b2cs.cell(5, 1, 'OE')
                ws_b2cs.cell(5, 2, '09-Uttar Pradesh')
                ws_b2cs.cell(5, 4, 18.0)
                ws_b2cs.cell(5, 5, t_taxable)
                ws_b2cs.cell(5, 6, 0.0)

        # 2. HSN Summary (b2c)
        hsn_dict = defaultdict(lambda: {'desc': set(), 'uqc': 'PCS-PIECES', 'qty': 0.0, 'total': 0.0, 'taxable': 0.0, 'cgst': 0.0, 'sgst': 0.0})
        for inv in b2c_invs:
            for it in inv['items']:
                h = it['hsn']
                hsn_dict[h]['desc'].add(it['desc'])
                hsn_dict[h]['qty'] += it['qty']
                hsn_dict[h]['total'] += it['total']
                hsn_dict[h]['taxable'] += it['taxable_value']
                hsn_dict[h]['cgst'] += it['cgst_amount']
                hsn_dict[h]['sgst'] += it['sgst_amount']

        if 'hsn(b2c)' in wb.sheetnames:
            ws_hsn = wb['hsn(b2c)']
            while ws_hsn.max_row > 4: ws_hsn.delete_rows(5)
            ws_hsn.cell(3, 1, len(hsn_dict))
            ws_hsn.cell(3, 5, round(sum(v['total'] for v in hsn_dict.values()), 2))
            ws_hsn.cell(3, 7, round(sum(v['taxable'] for v in hsn_dict.values()), 2))
            ws_hsn.cell(3, 9, round(sum(v['cgst'] for v in hsn_dict.values()), 2))
            ws_hsn.cell(3, 10, round(sum(v['sgst'] for v in hsn_dict.values()), 2))

            r = 5
            for hsn, data in sorted(hsn_dict.items()):
                ws_hsn.cell(r, 1, int(hsn))
                ws_hsn.cell(r, 2, ", ".join(sorted(data['desc'])))
                ws_hsn.cell(r, 3, data['uqc'])
                ws_hsn.cell(r, 4, data['qty'])
                ws_hsn.cell(r, 5, round(data['total'], 2))
                ws_hsn.cell(r, 6, 18.0)
                ws_hsn.cell(r, 7, round(data['taxable'], 2))
                ws_hsn.cell(r, 8, 0.0)
                ws_hsn.cell(r, 9, round(data['cgst'], 2))
                ws_hsn.cell(r, 10, round(data['sgst'], 2))
                ws_hsn.cell(r, 11, 0.0)
                r += 1

        # 3. Docs Sheet
        if 'docs' in wb.sheetnames:
            ws_docs = wb['docs']
            nums = [x['invoice_no'] for x in invoices if x['invoice_no']]
            ws_docs.cell(3, 4, len(nums))
            ws_docs.cell(3, 5, 0)
            ws_docs.cell(5, 1, 'Invoices for outward supply')
            ws_docs.cell(5, 2, min(nums) if nums else 1)
            ws_docs.cell(5, 3, max(nums) if nums else len(nums))
            ws_docs.cell(5, 4, len(nums))
            ws_docs.cell(5, 5, 0)

        # 4. Invoice Register Sheet
        if "Invoice_Register" in wb.sheetnames: del wb["Invoice_Register"]
        ws_reg = wb.create_sheet(title="Invoice_Register")
        headers = ['Invoice No', 'Date', 'Customer Name', 'Address', 'GSTIN', 'Item', 'HSN', 'Qty', 'Unit', 'Rate', 'Taxable Value', 'CGST', 'SGST', 'Total']
        ws_reg.append(headers)

        for c in range(1, len(headers)+1):
            ws_reg.cell(1, c).fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
            ws_reg.cell(1, c).font = Font(bold=True, color="FFFFFF")

        for inv in invoices:
            for it in inv['items']:
                ws_reg.append([
                    inv['invoice_no'], inv['invoice_date'], inv['customer_name'],
                    inv['address'], inv['cust_gstin'] or 'Unregistered', it['desc'], int(it['hsn']),
                    it['qty'], it['uqc'], it['rate'], it['taxable_value'],
                    it['cgst_amount'], it['sgst_amount'], it['total']
                ])

        # आउटपुट
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success(f"✅ कुल {len(invoices)} इनवॉइस सफलतापूर्वक प्रोसेस किए गए!")
        st.download_button(
            label="📥 Download Excel File (.xlsx)",
            data=output,
            file_name="GSTR1_Updated_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
