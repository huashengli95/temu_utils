# utilities/output_manager.py

import os
from datetime import datetime


def prepare_output_paths(base_dir=".", order_count=0):
    today = datetime.now().strftime("%Y-%m-%d")

    output_dir = os.path.join(
        base_dir,
        f"{today}_output"
    )
    os.makedirs(output_dir, exist_ok=True)

    pdf_path = os.path.join(
        output_dir,
        f"{today}_{order_count}.pdf"
    )

    excel_path = os.path.join(
        output_dir,
        f"{today}_{order_count}.xlsx"
    )

    return output_dir, pdf_path, excel_path
