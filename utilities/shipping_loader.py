# shipping_loader.py

import pandas as pd
from collections import defaultdict


def load_shipping_orders(shipping_excel_path: str):
    """
    return:
      order_no -> tracking_no
    """
    df = pd.read_excel(
        shipping_excel_path,
        dtype=str
    )

    df["订单号"] = df["订单号"].str.strip()
    df["跟踪单号"] = df["跟踪单号"].str.strip()

    return dict(zip(df["订单号"], df["跟踪单号"]))
