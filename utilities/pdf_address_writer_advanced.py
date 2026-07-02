# pdf_address_writer_advanced.py
# Advanced version: replace asterisks with actual address info

import re
from io import BytesIO
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter


def extract_postal_code_from_pdf(text: str) -> str:
    postal_pattern = re.compile(r"[A-Z]\d[A-Z]\s?\d[A-Z]\d")
    match = postal_pattern.search(text)
    if match:
        return match.group(0).replace(" ", "")
    return ""


def find_matching_order(text: str, orders: dict) -> tuple:
    pdf_postal_code = extract_postal_code_from_pdf(text)
    if not pdf_postal_code:
        return None, None
    
    postal_to_orders = {}
    for order_no, order_item in orders.items():
        postal = order_item.postal_code.upper().replace(" ", "")
        if postal not in postal_to_orders:
            postal_to_orders[postal] = []
        postal_to_orders[postal].append(order_item)
    
    pdf_postal_clean = pdf_postal_code.upper().replace(" ", "")
    if pdf_postal_clean in postal_to_orders:
        return postal_to_orders[pdf_postal_clean][0], pdf_postal_clean
    
    return None, None


def fill_address_info_to_pdf_advanced(
    input_pdf: str,
    output_pdf: str,
    orders: dict,
    debug: bool = False
):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    for page_idx, page in enumerate(reader.pages):
        text = page.extract_text() or ""
        
        if debug:
            print(f"\n{'='*60}")
            print(f"Processing page {page_idx + 1}")
            print(f"{'='*60}")
        
        matched_order, postal_code = find_matching_order(text, orders)
        
        if matched_order:
            if debug:
                print(f"[OK] Found matching order")
                print(f"   Recipient: {matched_order.recipient_name}")
                print(f"   Postal Code: {postal_code}")
                print(f"   Phone: {matched_order.phone}")
            
            phone = matched_order.phone
            if "-" in phone:
                phone = phone.split("-")[0]
            
            # Clear form fields if they exist
            if "/AcroForm" in page:
                page["/AcroForm"] = None
            
            packet = BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont("Helvetica", 9)
            
            # Draw white rectangles to cover the asterisks first
            # This ensures complete coverage of the original text
            can.setFillColor("white")
            can.setStrokeColor("white")
            
            # Cover recipient name area (first ****** line)
            can.rect(5, 336, 100, 9, fill=1, stroke=0)
            
            # Cover address line 1 area (second ****** line)
            can.rect(5, 326, 100, 10, fill=1, stroke=0)
            
            # Cover address line 2 area (third ****** line)
            can.rect(5, 317, 100, 10, fill=1, stroke=0)
            
            # Cover phone number area
            can.rect(89, 345, 70, 7, fill=1, stroke=0)
            
            # Now draw the actual text
            can.setFillColor("black")
            can.setFont("Helvetica", 8)
            
            # Recipient name (first ****** line)
            can.drawString(6, 335, matched_order.recipient_name)
            
            # Address line 1 (second ****** line)
            if matched_order.address1:
                can.drawString(6, 325, matched_order.address1)
            
            # Address line 2 (third ****** line)
            if matched_order.address2:
                can.drawString(6, 316, matched_order.address2)
            
            # Phone number (after "Tel. / Tél. :")
            can.drawString(90, 344, phone)
            
            can.save()
            packet.seek(0)
            
            overlay = PdfReader(packet)
            page.merge_page(overlay.pages[0])
            
            if debug:
                print(f"[OK] Filled info to PDF")
        else:
            if debug:
                print(f"[ERROR] No matching order found")
                postal = extract_postal_code_from_pdf(text)
                print(f"   PDF Postal Code: {postal}")
        
        writer.add_page(page)

    with open(output_pdf, "wb") as f:
        writer.write(f)
    
    if debug:
        print(f"\n[OK] PDF saved to: {output_pdf}")


def fill_address_info_to_pdf(
    input_pdf: str,
    output_pdf: str,
    orders: dict
):
    fill_address_info_to_pdf_advanced(input_pdf, output_pdf, orders, debug=False)
