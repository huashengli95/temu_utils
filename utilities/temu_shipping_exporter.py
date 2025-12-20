# temu_shipping_exporter.py

import pandas as pd


def export_temu_shipping_excel(
    tracking_to_suborders: dict,
    output_excel_path: str
):
    rows = []

    for tracking_no, suborders in tracking_to_suborders.items():
        for sub in suborders:
            rows.append({
                "订单号": sub.order_no,
                "子订单号": sub.sub_order_no,
                "商品件数": sub.quantity,
                "跟踪单号": tracking_no,
                "物流承运商": "FedEx",
                "发货仓库名称": "Deerchase court"
            })

    pd.DataFrame(rows).to_excel(
        output_excel_path,
        index=False
    )
