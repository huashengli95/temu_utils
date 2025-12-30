# main.py

import sys
from utilities.temu_loader import load_temu_orders
from utilities.shipping_loader import load_shipping_orders
from utilities.pdf_label_writer import add_sku_info_to_pdf
from utilities.temu_shipping_exporter import export_temu_shipping_excel
from utilities.output_manager import prepare_output_paths
from pathlib import Path


def find_single_file(base_dir, *, suffix=None, name_contains=None):
    base_dir = Path(base_dir)

    files = [
        f for f in base_dir.iterdir()
        if f.is_file()
        and (suffix is None or f.suffix == suffix)
        and (name_contains is None or name_contains in f.name)
    ]

    if not files:
        raise FileNotFoundError(
            f"No file found in {base_dir} "
            f"(suffix={suffix}, name_contains={name_contains})"
        )

    if len(files) > 1:
        raise RuntimeError(
            f"Multiple files found in {base_dir}: "
            f"{[f.name for f in files]}"
        )

    return files[0]


def main():
    shipping_pdf = find_single_file(
        ".",
        suffix=".pdf"
    )

    shipping_excel = find_single_file(
        "./output",
        suffix=".xlsx",
        name_contains="to_agent"
    )

    temu_excel = find_single_file(
        ".",
        suffix=".xlsx",
        name_contains="订单导出"
    )

    # 1️⃣ 加载 Temu 订单
    orders = load_temu_orders(temu_excel)

    # 2️⃣ 订单号 -> 跟踪单号
    order_to_tracking = load_shipping_orders(shipping_excel)

    # 3️⃣ tracking_no -> SubOrderItem[]
    tracking_to_suborders = {}

    for order_no, tracking_no in order_to_tracking.items():
        order = orders.get(order_no)
        if not order:
            continue

        for sub in order.sub_orders:
            sub.order_no = order_no

        tracking_to_suborders.setdefault(
            tracking_no, []
        ).extend(order.sub_orders)

    order_count = len(tracking_to_suborders)

    # 4️⃣ 准备输出路径
    _, pdf_path, excel_path = prepare_output_paths(
        base_dir=".",
        order_count=order_count
    )

    # 5️⃣ 写 PDF
    add_sku_info_to_pdf(
        shipping_pdf,
        pdf_path,
        tracking_to_suborders
    )

    # 6️⃣ 导出 Temu 运单 Excel
    export_temu_shipping_excel(
        tracking_to_suborders,
        excel_path
    )

    print(f"✅ PDF 输出: {pdf_path}")
    print(f"✅ Excel 输出: {excel_path}")


if __name__ == "__main__":
    main()