#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
生成可回填 Temu 的 Excel 文件

用法:
    python generate_temu_backfill_excel.py <订单导出.xlsx> <输出文件.xlsx>

说明:
    - 从 processed_data.xlsx 中读取订单号和运单号的映射
    - 从订单导出 Excel 中读取订单详情
    - 生成可回填 Temu 的 Excel 文件
"""

import sys
import io
import pandas as pd
from pathlib import Path
from utilities.temu_loader import load_temu_orders
from utilities.temu_shipping_exporter import export_temu_shipping_excel

# 设置输出编码为 UTF-8
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


def load_processed_data(processed_data_path: str):
    """
    从 processed_data.xlsx 中读取订单号和运单号的映射
    return:
        {
            order_no: tracking_no
        }
    """
    df = pd.read_excel(processed_data_path)
    
    # 提取订单号和运单号列
    order_to_tracking = {}
    
    for _, row in df.iterrows():
        order_no = str(row.get("订单号(必填)", "")).strip()
        tracking_no = str(row.get("运单号", "")).strip()
        
        if order_no and tracking_no:
            order_to_tracking[order_no] = tracking_no
    
    return order_to_tracking


def main():
    if len(sys.argv) < 3:
        print("用法: python generate_temu_backfill_excel.py <订单导出.xlsx> <processed_data.xlsx> [输出文件.xlsx]")
        print()
        print("说明:")
        print("  - 从 processed_data.xlsx 中读取订单号和运单号的映射")
        print("  - 从订单导出 Excel 中读取订单详情")
        print("  - 生成可回填 Temu 的 Excel 文件")
        sys.exit(1)
    
    temu_excel = sys.argv[1]
    processed_data_path = sys.argv[2]
    
    # 确定输出文件路径
    if len(sys.argv) >= 4:
        output_excel = sys.argv[3]
    else:
        # 默认输出到 output 目录
        output_dir = Path("output")
        output_dir.mkdir(exist_ok=True)
        output_excel = output_dir / "temu_backfill.xlsx"
    
    if not Path(processed_data_path).exists():
        print(f"错误: 找不到 {processed_data_path}")
        sys.exit(1)
    
    if not Path(temu_excel).exists():
        print(f"错误: 找不到 {temu_excel}")
        sys.exit(1)
    
    print(f"读取订单数据: {temu_excel}")
    print(f"读取运单映射: {processed_data_path}")
    
    # 加载 Temu 订单
    orders = load_temu_orders(temu_excel)
    print(f"加载了 {len(orders)} 个订单")
    
    # 从 processed_data.xlsx 读取订单号 -> 运单号映射
    order_to_tracking = load_processed_data(str(processed_data_path))
    print(f"加载了 {len(order_to_tracking)} 个订单-运单映射")
    
    # 构建 tracking_no -> SubOrderItem[] 映射
    tracking_to_suborders = {}
    
    for order_no, tracking_no in order_to_tracking.items():
        order = orders.get(order_no)
        if not order:
            print(f"警告: 订单 {order_no} 在订单导出中找不到")
            continue
        
        for sub in order.sub_orders:
            sub.order_no = order_no
        
        tracking_to_suborders.setdefault(
            tracking_no, []
        ).extend(order.sub_orders)
    
    print(f"构建了 {len(tracking_to_suborders)} 个运单的子订单映射")
    
    # 导出 Temu 运单 Excel
    export_temu_shipping_excel(
        tracking_to_suborders,
        str(output_excel)
    )
    
    print(f"Excel 已输出: {output_excel}")


if __name__ == "__main__":
    main()
