# generatePDF_v2.py
# 新功能：根据 Temu Excel 中的姓名和邮编匹配面单，填充地址和电话信息

import sys
from utilities.temu_loader import load_temu_orders
from utilities.pdf_address_writer_advanced import fill_address_info_to_pdf_advanced
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
    # 查找输入文件
    shipping_pdf = find_single_file(
        ".",
        suffix=".pdf"
    )

    temu_excel = find_single_file(
        ".",
        suffix=".xlsx",
        name_contains="订单导出"
    )

    # 1. 加载 Temu 订单（包含地址信息）
    print("[*] Loading Temu orders...")
    orders = load_temu_orders(temu_excel)

    print(f"[OK] Loaded {len(orders)} orders\n")
    
    # 显示订单信息用于调试
    print("[*] Order preview (first 3):")
    for i, (order_no, order) in enumerate(list(orders.items())[:3], 1):
        print(f"  {i}. {order_no}")
        print(f"     Recipient: {order.recipient_name}")
        print(f"     Postal Code: {order.postal_code}")
        print(f"     Phone: {order.phone}")
        print(f"     Address: {order.address1}")
    print()

    # 2. 准备输出路径
    order_count = len(orders)
    _, pdf_path, _ = prepare_output_paths(
        base_dir=".",
        order_count=order_count
    )

    # 3. 填充地址信息到 PDF
    print("[*] Filling PDF labels...")
    fill_address_info_to_pdf_advanced(
        shipping_pdf,
        pdf_path,
        orders,
        debug=True
    )

    print(f"\n[OK] PDF output: {pdf_path}")


if __name__ == "__main__":
    main()
