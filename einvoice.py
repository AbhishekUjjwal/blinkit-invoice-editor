import streamlit as st
import fitz  # PyMuPDF
import re
import io
import json
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="e-Invoice Gateway & Excel Suite", page_icon="🧾", layout="centered")

st.markdown("""
    <div style="text-align: center; padding: 10px 0 20px 0;">
        <h2 style="color: #FFFFFF; margin-bottom: 4px;">⚡ Blinkit e-Invoice Gateway</h2>
        <p style="color: #94a3b8; font-size: 14px;">Government Portal Ready JSON & Official NIC Bulk Excel Generator</p>
    </div>
""", unsafe_allow_html=True)

uploaded_invoices = st.file_uploader("Upload Blinkit Invoices (PDF)", type=["pdf"], accept_multiple_files=True)

def extract_metadata(full_text):
    inv_match = re.search(r"Invoice\s*No\s*[:\s]*([A-Za-z0-9\-_]+)", full_text, re.IGNORECASE)
    invoice_no = inv_match.group(1).strip() if inv_match else "INV-001"

    date_match = re.search(r"Invoice\s*Date\s*[:\s]*([A-Za-z0-9,\s\.\-\/]+?)(?=\n|Ship\s*Date|$)", full_text, re.IGNORECASE)
    raw_date = date_match.group(1).strip() if date_match else "Date"

    std_date_for_json = datetime.now().strftime("%d/%m/%Y")
    for fmt in ("%b %d, %Y", "%B %d, %Y", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            d_obj = datetime.strptime(raw_date, fmt)
            std_date_for_json = d_obj.strftime("%d/%m/%Y")
            break
        except ValueError:
            pass

    return invoice_no, std_date_for_json

def parse_pdf_data(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    full_text = ""
    for p in doc:
        full_text += p.get_text() + "\n"

    invoice_no, std_date = extract_metadata(full_text)

    gstins = re.findall(r"\b\d{2}[A-Z]{5}\d{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}\b", full_text)
    seller_gstin = gstins[0] if len(gstins) > 0 else "09AAFCG9846E1Z9"
    buyer_gstin = gstins[1] if len(gstins) > 1 else seller_gstin

    pincodes = re.findall(r"\b[1-9][0-9]{5}\b", full_text)
    seller_pin = int(pincodes[0]) if len(pincodes) > 0 else 226401
    buyer_pin = int(pincodes[1]) if len(pincodes) > 1 else seller_pin

    seller_state_code = seller_gstin[:2]
    buyer_state_code = buyer_gstin[:2]

    buyer_name = "BLINK COMMERCE PRIVATE LIMITED"

    line_items = []
    for page in doc:
        words = page.get_text("words")
        qty_box = None
        price_box = None

        for w in words:
            w_txt = w[4].strip().lower()
            if w_txt == "qty":
                qty_box = (w[0] - 15, w[2] + 25)
            elif w_txt == "price" or "unit" in w_txt:
                if not price_box:
                    price_box = (w[0] - 15, w[2] + 40)

        if not qty_box:
            qty_box = (320, 390)
        if not price_box:
            price_box = (390, 480)

        product_rows = []
        for w in words:
            text = w[4].upper()
            if "DISPO" in text:
                product_rows.append({"type": "DISPO", "factor": 50, "y": (w[1] + w[3]) / 2, "desc": "DGNW50 Dispo Guard Face Mask"})
            elif "COMFIT" in text:
                product_rows.append({"type": "COMFIT", "factor": 25, "y": (w[1] + w[3]) / 2, "desc": "C3DFM Comfit 3D Face Mask"})

        for prod in product_rows:
            factor = prod["factor"]
            row_y = prod["y"]
            row_words = [w for w in words if abs(((w[1] + w[3]) / 2) - row_y) <= 25]

            final_q = None
            final_p = None

            for w in row_words:
                val = w[4].replace(",", "").strip()
                if qty_box[0] <= w[0] <= qty_box[1]:
                    if val.isdigit() and len(val) != 8:
                        orig_q = int(val)
                        if orig_q >= factor:
                            final_q = orig_q // factor

                if price_box[0] <= w[0] <= price_box[1]:
                    if re.match(r"^\d+\.\d{2}$", val):
                        orig_p = float(val)
                        if orig_p > 0.00:
                            final_p = round(orig_p * factor, 2)

            if final_q and final_p:
                line_items.append({
                    "desc": prod["desc"],
                    "hsn": "63079091",
                    "qty": final_q,
                    "unit_price": final_p,
                    "gst_rate": 5.0
                })

    doc.close()

    return {
        "invoice_no": invoice_no,
        "doc_date": std_date,
        "seller_gstin": seller_gstin,
        "buyer_gstin": buyer_gstin,
        "seller_pin": seller_pin,
        "buyer_pin": buyer_pin,
        "seller_state_code": seller_state_code,
        "buyer_state_code": buyer_state_code,
        "buyer_name": buyer_name,
        "line_items": line_items
    }

def build_einvoice_json_object(data):
    seller_state_code = data["seller_state_code"]
    buyer_state_code = data["buyer_state_code"]

    item_list = []
    tot_taxable = 0.0
    tot_cgst = 0.0
    tot_sgst = 0.0
    tot_igst = 0.0

    for idx, item in enumerate(data["line_items"], 1):
        qty = item["qty"]
        unit_price = item["unit_price"]
        taxable_amt = round(qty * unit_price, 2)
        tot_taxable += taxable_amt

        gst_rate = item.get("gst_rate", 5.0)
        is_interstate = (seller_state_code != buyer_state_code)

        if is_interstate:
            igst_amt = round((taxable_amt * gst_rate) / 100, 2)
            cgst_amt = 0.0
            sgst_amt = 0.0
            tot_igst += igst_amt
        else:
            cgst_amt = round((taxable_amt * (gst_rate / 2)) / 100, 2)
            sgst_amt = round((taxable_amt * (gst_rate / 2)) / 100, 2)
            igst_amt = 0.0
            tot_cgst += cgst_amt
            tot_sgst += sgst_amt

        tot_item_val = round(taxable_amt + cgst_amt + sgst_amt + igst_amt, 2)

        item_list.append({
            "SlNo": str(idx),
            "PrdDesc": item["desc"][:100],
            "IsServc": "N",
            "HsnCd": item["hsn"],
            "Qty": qty,
            "Unit": "BOX",
            "UnitPrice": unit_price,
            "TotAmt": taxable_amt,
            "Discount": 0.0,
            "AssAmt": taxable_amt,
            "GstRt": gst_rate,
            "IgstAmt": igst_amt,
            "CgstAmt": cgst_amt,
            "SgstAmt": sgst_amt,
            "CesRt": 0.0,
            "CesAmt": 0.0,
            "TotItemVal": tot_item_val
        })

    total_inv_val = round(tot_taxable + tot_cgst + tot_sgst + tot_igst, 2)

    return {
        "Version": "1.03",
        "TranDtls": {
            "TaxSch": "GST",
            "SupTyp": "B2B",
            "RegRev": "N",
            "EcmGstin": None,
            "IgstOnIntra": "N"
        },
        "DocDtls": {
            "Typ": "INV",
            "No": data["invoice_no"],
            "Dt": data["doc_date"]
        },
        "SellerDtls": {
            "Gstin": data["seller_gstin"],
            "LglNm": "SELLER TRADING CO",
            "TrdNm": "SELLER TRADING CO",
            "Addr1": "Warehouse Address",
            "Loc": "City",
            "Pin": data["seller_pin"],
            "Stcd": seller_state_code
        },
        "BuyerDtls": {
            "Gstin": data["buyer_gstin"],
            "LglNm": data["buyer_name"],
            "TrdNm": data["buyer_name"],
            "Pos": buyer_state_code,
            "Addr1": "Warehouse Facility, Junabganj Road",
            "Loc": "Lucknow",
            "Pin": data["buyer_pin"],
            "Stcd": buyer_state_code
        },
        "ItemList": item_list,
        "ValDtls": {
            "AssVal": round(tot_taxable, 2),
            "CgstVal": round(tot_cgst, 2),
            "SgstVal": round(tot_sgst, 2),
            "IgstVal": round(tot_igst, 2),
            "CesVal": 0.0,
            "StCesVal": 0.0,
            "RndOffAmt": 0.0,
            "TotInvVal": total_inv_val
        }
    }

def generate_official_nic_excel(data_list):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "eInvoice"
    ws.views.sheetView[0].showGridLines = True

    # Styling colors
    sky_blue = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
    light_blue = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    peach = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid")
    green = PatternFill(start_color="D8E4BC", end_color="D8E4BC", fill_type="solid")

    font_title = Font(name="Arial", size=18, bold=True)
    font_button = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    font_sec = Font(name="Arial", size=10, bold=True)
    font_col = Font(name="Arial", size=9, bold=True)
    font_data = Font(name="Arial", size=9)

    thin_border = Border(
        left=Side(style='thin', color='B0B0B0'),
        right=Side(style='thin', color='B0B0B0'),
        top=Side(style='thin', color='B0B0B0'),
        bottom=Side(style='thin', color='B0B0B0')
    )

    ws["G1"] = "E-Invoice System"
    ws["G1"].font = font_title
    ws["G1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("L1:N2")
    ws["L1"] = "Validate"
    ws["L1"].fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    ws["L1"].font = font_button
    ws["L1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("P1:R2")
    ws["P1"] = "Prepare JSON"
    ws["P1"].fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    ws["P1"].font = font_button
    ws["P1"].alignment = Alignment(horizontal="center", vertical="center")

    ws["T2"] = "* Indicates Mandatory Fields"
    ws["T2"].font = Font(name="Arial", size=8, italic=True, color="555555")

    sections = [
        ("Supply Details", 1, 4, light_blue),
        ("Document Details", 5, 7, peach),
        ("Buyer Details", 8, 18, light_blue),
        ("Dispatch Details", 19, 24, peach),
        ("Item Details", 25, 38, green),
        ("Invoice Value Details", 39, 44, sky_blue)
    ]

    for name, sc, ec, fill in sections:
        ws.merge_cells(start_row=3, start_column=sc, end_row=3, end_column=ec)
        c = ws.cell(row=3, column=sc, value=name)
        c.fill = fill
        c.font = font_sec
        c.alignment = Alignment(horizontal="center", vertical="center")
        for i in range(sc, ec + 1):
            ws.cell(row=3, column=i).border = thin_border

    columns = [
        "Supply Type *", "Reverse Charge", "e-Comm GSTIN", "Igst On Intra",
        "Document Type *", "Document Number *", "Document Date (DD/MM/YYYY) *",
        "Buyer GSTIN *", "Buyer Legal Name *", "Buyer Trade Name", "Buyer POS *", 
        "Buyer Addr1 *", "Buyer Addr2", "Buyer Location *", "Buyer Pin Code *", 
        "Buyer State *", "Buyer Phone Number", "Buyer Email Id",
        "Dispatch Name", "Dispatch Addr1", "Dispatch Addr2", "Dispatch Location", "Dispatch Pin Code", "Dispatch State",
        "Sl. No *", "Product Description *", "Is Service *", "HSN Code *", "Quantity *", "Unit *", "Unit Price *", 
        "Gross Amount", "Taxable Value *", "GST Rate (%) *", "IGST Amount", "CGST Amount", "SGST Amount", "Total Item Value *",
        "Total Taxable Value *", "Total CGST Amount", "Total SGST Amount", "Total IGST Amount", "Round Off Amount", "Total Invoice Value *"
    ]

    for c_idx, col_name in enumerate(columns, 1):
        cell = ws.cell(row=4, column=c_idx, value=col_name)
        if c_idx <= 4:
            cell.fill = light_blue
        elif c_idx <= 7:
            cell.fill = peach
        elif c_idx <= 18:
            cell.fill = light_blue
        elif c_idx <= 24:
            cell.fill = peach
        elif c_idx <= 38:
            cell.fill = green
        else:
            cell.fill = sky_blue
        cell.font = font_col
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    ws.row_dimensions[4].height = 40

    curr_row = 5
    for inv in data_list:
        seller_state_code = inv["seller_state_code"]
        buyer_state_code = inv["buyer_state_code"]

        tot_taxable = sum([round(it["qty"] * it["unit_price"], 2) for it in inv["line_items"]])
        is_interstate = (seller_state_code != buyer_state_code)

        if is_interstate:
            tot_igst = round(tot_taxable * 0.05, 2)
            tot_cgst = 0.0
            tot_sgst = 0.0
        else:
            tot_igst = 0.0
            tot_cgst = round(tot_taxable * 0.025, 2)
            tot_sgst = round(tot_taxable * 0.025, 2)

        tot_inv_val = round(tot_taxable + tot_cgst + tot_sgst + tot_igst, 2)

        for s_no, it in enumerate(inv["line_items"], 1):
            qty = it["qty"]
            price = it["unit_price"]
            taxable = round(qty * price, 2)
            gst_rate = it.get("gst_rate", 5.0)

            if is_interstate:
                igst = round(taxable * (gst_rate / 100), 2)
                cgst = 0.0
                sgst = 0.0
            else:
                cgst = round(taxable * (gst_rate / 200), 2)
                sgst = round(taxable * (gst_rate / 200), 2)
                igst = 0.0

            item_val = round(taxable + cgst + sgst + igst, 2)

            row_data = [
                "B2B", "N", "", "N",
                "Tax Invoice", inv["invoice_no"], inv["doc_date"],
                inv["buyer_gstin"], inv["buyer_name"], inv["buyer_name"], buyer_state_code,
                "Warehouse Facility, Junabganj Road", "", "Lucknow", inv["buyer_pin"],
                "UTTAR PRADESH", "9139396200", "billing@blinkit.com",
                "Warehouse Hub", "Industrial Area", "", "Lucknow", inv["seller_pin"], seller_state_code,
                s_no, it["desc"], "N", it["hsn"], qty, "BOX", price,
                taxable, taxable, gst_rate, igst, cgst, sgst, item_val,
                tot_taxable, tot_cgst, tot_sgst, tot_igst, 0.00, tot_inv_val
            ]

            for c_idx, val in enumerate(row_data, 1):
                cell = ws.cell(row=curr_row, column=c_idx, value=val)
                cell.font = font_data
                cell.border = thin_border
                if c_idx in [1, 2, 4, 5, 7, 8, 11, 15, 16, 23, 24, 25, 27, 28, 30]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c_idx == 29:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = "#,##0"
                elif c_idx in [31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44]:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = "#,##0.00"
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

            curr_row += 1

    tabs = ["Welcome", "Profile", "Master Codes", "Sample Invoice", "Format A,B,C,D", "Schema", "Validation", "Calculations", "FAQs"]
    for t in tabs:
        d_ws = wb.create_sheet(title=t)
        d_ws.sheet_view.showGridLines = True
        d_ws["A1"] = f"{t} - NIC e-Invoice Offline Utility"
        d_ws["A1"].font = Font(size=14, bold=True, color="1F497D")

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row < 3:
                continue
            v_str = str(cell.value or '')
            if len(v_str) > max_len:
                max_len = len(v_str)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

    out_io = io.BytesIO()
    wb.save(out_io)
    out_io.seek(0)
    return out_io

if uploaded_invoices:
    parsed_invoices = []
    for file in uploaded_invoices:
        try:
            data = parse_pdf_data(file.read())
            parsed_invoices.append(data)
        except Exception as e:
            st.error(f"Error reading {file.name}: {str(e)}")

    if parsed_invoices:
        st.success(f"Successfully processed {len(parsed_invoices)} Invoices!")
        
        # Prepare Excel & JSON outputs
        excel_buffer = generate_official_nic_excel(parsed_invoices)

        if len(parsed_invoices) == 1:
            inv = parsed_invoices[0]
            json_payload = build_einvoice_json_object(inv)
            inv_no = inv["invoice_no"]

            c1, c2 = st.columns(2)
            with c1:
                st.download_button(
                    label=f"🧾 Download e-Invoice JSON",
                    data=json.dumps(json_payload, indent=4),
                    file_name=f"{inv_no}_eInvoice.json",
                    mime="application/json"
                )
            with c2:
                st.download_button(
                    label=f"📊 Download NIC Bulk Excel (.xlsx)",
                    data=excel_buffer,
                    file_name=f"{inv_no}_NIC_eInvoice.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            bulk_json = [build_einvoice_json_object(inv) for inv in parsed_invoices]
            date_str = datetime.now().strftime('%d-%m-%Y')

            c1, c2 = st.columns(2)
            with c1:
                st.download_button(
                    label=f"🧾 Download Bulk e-Invoice JSON ({len(bulk_json)} Invoices)",
                    data=json.dumps(bulk_json, indent=4),
                    file_name=f"Bulk_eInvoice_NIC_{date_str}.json",
                    mime="application/json"
                )
            with c2:
                st.download_button(
                    label=f"📊 Download Bulk NIC Excel ({len(parsed_invoices)} Invoices)",
                    data=excel_buffer,
                    file_name=f"Bulk_NIC_eInvoice_{date_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
