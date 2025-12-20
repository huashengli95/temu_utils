# temu_loader.py

import pandas as pd
from typing import Dict
from collections import defaultdict
from utilities.order_item import SubOrderItem, OrderItem


def load_temu_orders(temu_order_excel_path: str) -> Dict[str, OrderItem]:
    """
    从 Temu 导出的订单 Excel 构建 OrderItem 映射
    return:
        {
            order_no: OrderItem(...)
        }
    """
    df = pd.read_excel(temu_order_excel_path, dtype=str)

    df["订单号"] = df["订单号"].str.strip()
    df["子订单号"] = df["子订单号"].str.strip()
    df["SKU货号"] = df["SKU货号"].str.strip()
    df["应履约件数"] = df["应履约件数"].astype(int)

    orders: Dict[str, OrderItem] = {}

    for _, row in df.iterrows():
        order_no = row["订单号"]

        if order_no not in orders:
            orders[order_no] = OrderItem(order_no)

        sub = SubOrderItem(
            sub_order_no=row["子订单号"],
            sku=row["SKU货号"],
            quantity=row["应履约件数"]
        )

        orders[order_no].add_sub_order(sub)

    return orders
