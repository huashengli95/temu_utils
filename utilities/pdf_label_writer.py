# pdf_label_writer.py

import re
from io import BytesIO
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

FEDEX_PATTERN = re.compile(r"\b\d{4}\s\d{4}\s\d{4}\b")


def extract_tracking_numbers(pdf_path: str):
    reader = PdfReader(pdf_path)
    results = set()

    for page in reader.pages:
        text = page.extract_text() or ""
        for m in FEDEX_PATTERN.findall(text):
            results.add(m.replace(" ", ""))

    return results


def add_sku_info_to_pdf(
    input_pdf: str,
    output_pdf: str,
    tracking_to_suborders: dict
):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    for page in reader.pages:
        text = page.extract_text() or ""
        matches = FEDEX_PATTERN.findall(text)
        tracking_numbers = [m.replace(" ", "") for m in matches]

        lines = []
        for tn in tracking_numbers:
            for sub in tracking_to_suborders.get(tn, []):
                lines.append(f"{sub.quantity}*{sub.sku}")

        if lines:
            packet = BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont("Helvetica", 8)

            y = 410
            for line in lines:
                can.drawString(225, y, line)
                y += 9

            can.save()
            packet.seek(0)

            overlay = PdfReader(packet)
            page.merge_page(overlay.pages[0])

        writer.add_page(page)

    with open(output_pdf, "wb") as f:
        writer.write(f)
