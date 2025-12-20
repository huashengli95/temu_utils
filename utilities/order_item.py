# order_item.py

from typing import Dict, List


class SubOrderItem:
    """
    子订单（SKU 级）
    """
    def __init__(self, sub_order_no: str, sku: str, quantity: int):
        self.sub_order_no = sub_order_no
        self.sku = sku
        self.quantity = quantity

    def __repr__(self):
        return (
            f"SubOrderItem(sub_order_no={self.sub_order_no}, "
            f"sku={self.sku}, quantity={self.quantity})"
        )


class OrderItem:
    """
    订单（订单号级）
    """
    def __init__(self, order_no: str):
        self.order_no = order_no
        self.sub_orders: List[SubOrderItem] = []

    def add_sub_order(self, sub_order: SubOrderItem):
        self.sub_orders.append(sub_order)

    def total_quantity(self) -> int:
        """
        该订单下所有 SKU 的总件数
        """
        return sum(s.quantity for s in self.sub_orders)

    def __repr__(self):
        return (
            f"OrderItem(order_no={self.order_no}, "
            f"sub_orders={self.sub_orders})"
        )
