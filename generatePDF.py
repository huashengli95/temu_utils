# main.py

import sys
from utilities.temu_loader import load_temu_orders
from utilities.shipping_loader import load_shipping_orders
from utilities.pdf_label_writer import add_sku_info_to_pdf
from utilities.temu_shipping_exporter import export_temu_shipping_excel
from utilities.output_manager import prepare_output_paths


def main():
    shipping_pdf = sys.argv[1]
    shipping_excel = sys.argv[2]
    temu_excel = sys.argv[3]

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