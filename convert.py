import pandas as pd
import sys
import random
import re
import os
import shutil
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from pathlib import Path
import json

# Load settings from ./config
STATE_MAP = {
    "NL": "NF",
    "QC": "PQ",
}

def normalize_state_for_new_template(state):
    if not state:
        return ""

    s = str(state).strip().upper()
    return STATE_MAP.get(s, s)
def clean_temu_phone(phone):
    if not phone:
        return ""

    s = str(phone).strip()
    return s.split("-", 1)[0]

def _load_json_config(filename):
    base_dir = Path(__file__).resolve().parent
    config_file = base_dir / "config" / filename
    with open(config_file, "r", encoding="utf-8") as f:
        return json.load(f)

SENDER_PROFILES = _load_json_config("sender_profiles.json")
SKU_PACKAGE_SPECS = _load_json_config("sku_package_specs.json")
#SKU_PACKAGE_SPECS = _load_json_config("sku_package_specs_inch.json")


def merge_skus_with_qty(group):
    """
    group: 同一个订单号的 dataframe
    返回:
      sku_display_str: "G300 × 2, G600 × 1"
      need_highlight: bool
    """
    sku_qty = {}

    for _, row in group.iterrows():
        sku = str(row["SKU货号"]).strip()
        qty = int(row["应履约件数"])

        sku_qty[sku] = sku_qty.get(sku, 0) + qty

    parts = []
    need_highlight = False

    # 多 SKU
    if len(sku_qty) > 1:
        need_highlight = True

    for sku, qty in sku_qty.items():
        parts.append(f"{sku} × {qty}")
        if qty != 1:
            need_highlight = True

    return ", ".join(parts), need_highlight

def build_output_filename(shipitem_count, ext=".xlsx"):
    """
    生成: YYYYMMDD_<count>.xlsx
    """
    date_str = datetime.now().strftime("%Y%m%d")
    return f"{date_str}_{shipitem_count}_合并{ext}"

def copy_template_with_custom_name(template_path, target_dir, shipitem_count):
    """
    拷贝模板到目录，并按规则重命名
    """
    new_filename = build_output_filename(shipitem_count)
    target_path = os.path.join(target_dir, new_filename)

    shutil.copy2(template_path, target_path)
    return target_path

def create_timestamp_dir(base_dir="output"):
    """
    创建 output/YYYYMMDD_HHMM 目录
    返回创建好的目录路径
    """
    os.makedirs(base_dir, exist_ok=True)
    return base_dir

def copy_template_to_dir(template_path, target_dir):
    """
    把模板拷贝到目标目录
    返回新模板路径
    """
    template_name = os.path.basename(template_path)
    target_path = os.path.join(target_dir, template_name)

    shutil.copy2(template_path, target_path)
    return target_path

def merge_skus(sku_series):
    """
    sku_series: pandas Series
    返回: 去重后的 SKU 字符串
    """
    skus = (
        sku_series
        .dropna()
        .astype(str)
        .str.strip()
        .unique()
    )
    return ", ".join(skus)

def merge_address(*parts, sep=", "):
    clean_parts = []

    for p in parts:
        # 1. None
        if p is None:
            continue

        # 2. NaN
        if isinstance(p, float) and pd.isna(p):
            continue

        s = str(p).strip()

        # 3. 空字符串
        if not s:
            continue

        # 4. 仅由 '-' 组成的占位符（-, --, --- 等）
        if re.fullmatch(r"-+", s):
            continue

        clean_parts.append(s)

    return sep.join(clean_parts)

def clear_template_except_header(template_path, sheet_name=None):
    wb = load_workbook(template_path)
    ws = wb[sheet_name] if sheet_name else wb.active

    # 删除第 2 行到最后一行
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    wb.save(template_path)

def fill_template(template_path, ship_items, dict_builder, sheet_name=None):
    wb = load_workbook(template_path)
    ws = wb[sheet_name] if sheet_name else wb.active

    headers = [cell.value for cell in ws[1]]
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

    start_row = 2
    for row_idx, item in enumerate(ship_items, start=start_row):
        item_dict = dict_builder(item)

        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=item_dict.get(header, ""))

            if item.need_highlight:
                cell.fill = green_fill

    wb.save(template_path)

def get_random_sender():
    return random.choice(SENDER_PROFILES)

class PackageSpec:
    def __init__(self, weight, length, width, height, package_count=1):
        self.weight = weight
        self.length = length
        self.width = width
        self.height = height
        self.package_count = package_count

DEFAULT_PACKAGE_SPEC = PackageSpec(
    weight=0.5,
    length=10,
    width=10,
    height=5,
    package_count=1
)

def get_package_spec_by_sku(sku):
    """
    sku 可能的形式:
      - G300
      - G300 × 2
      - G300 x2
      - G300 X 2
    这里只提取 SKU 本身用于查表
    """

    if not sku:
        print("⚠ SKU 为空，使用默认包装规格")
        return DEFAULT_PACKAGE_SPEC

    # 转成字符串并去空格
    sku_str = str(sku).strip()

    # 提取 SKU 本体（取第一个非空、非乘号的部分）
    # 例如 "G300 × 2" -> "G300"
    base_sku = re.split(r"\s*[xX×]\s*\d+", sku_str)[0].strip()

    spec = SKU_PACKAGE_SPECS.get(base_sku)

    if not spec:
        print(f"⚠ 未找到 SKU 包装规格: {base_sku}（原始值: {sku_str}），使用默认值")
        return DEFAULT_PACKAGE_SPEC

    return PackageSpec(**spec)

# 1. 定义 ShipItem 类
class ShipItem:
    def __init__(self, customer_order, sku, tracking_number,
                 receiver_name, receiver_phone, receiver_address,
                 receiver_zip, receiver_city, receiver_state, receiver_country,
                 need_highlight=False):
        self.customer_order = customer_order
        self.sku = sku
        self.tracking_number = tracking_number

        self.receiver_name = receiver_name
        self.receiver_phone = receiver_phone
        self.receiver_address = receiver_address
        self.receiver_zip = receiver_zip
        self.receiver_city = receiver_city
        self.receiver_state = receiver_state
        self.receiver_country = receiver_country
        self.need_highlight = need_highlight

        first_sku = sku.split(",")[0].strip()
        self.package_spec = get_package_spec_by_sku(first_sku)

        # 在对象内部随机绑定一个发件人
        self.sender = get_random_sender()

    def to_dict(self):
        return {
            "客户单号": self.customer_order,
            "SKU": self.sku,
            "快递单号": self.tracking_number,

            "发件人姓名": self.sender["name"],
            "发件人电话": self.sender["phone"],
            "发件人地址": self.sender["address"],
            "发件人邮编": self.sender["zip"],
            "发件人城市": self.sender["city"],
            "发件人州/省": self.sender["state"],
            "国家": self.sender["country"],
            "邮箱": self.sender["email"],

            "收件人名字": self.receiver_name,
            "收件人电话": self.receiver_phone,
            "收件人地址": self.receiver_address,
            "收件人邮编": self.receiver_zip,
            "收件人城市": self.receiver_city,
            "收件人州/省": self.receiver_state,
            "国家": self.receiver_country,

            "包裹数量": self.package_spec.package_count,
            "包裹重量": self.package_spec.weight,
            "重量单位KG": "KG",
            "长/厘米": self.package_spec.length,
            "宽/厘米": self.package_spec.width,
            "高/厘米": self.package_spec.height
        }
    
    def to_fedex_dict(self):
        return {
            "serviceType": "",
            "shipmentType": "OUTBOUND",

            "senderContactName": self.sender["name"],
            "senderContactNumber": self.sender["phone"],
            "senderLine1": self.sender["address"],
            "senderPostcode": self.sender["zip"],
            "senderCity": self.sender["city"],
            "senderState": normalize_state_for_new_template(self.sender["state"]),
            "senderCountry": "CA",
            "senderEmail": self.sender["email"],

            "recipientContactName": self.receiver_name,
            "recipientContactNumber": clean_temu_phone(self.receiver_phone),
            "recipientLine1": self.receiver_address,
            "recipientLine2": "",
            "recipientPostcode": self.receiver_zip,
            "recipientCity": self.receiver_city,
            "recipientState": normalize_state_for_new_template(self.receiver_state),
            "recipientCountry": "CA",

            "numberOfPackages": self.package_spec.package_count,
            "packageWeight": self.package_spec.weight,
            "weightUnits": "KGS",

            "length": self.package_spec.length,
            "width": self.package_spec.width,
            "height": self.package_spec.height,

            "packageType": "YOUR_PACKAGING",
            "currencyType": "CAD",
            "Package contents": self.sku,
        }

# 1. 读取模板 Excel
BASE_DIR = Path(__file__).resolve().parent
template_excel = BASE_DIR / "fedexTemplate.xlsx"
fedex_template_excel = BASE_DIR / "newShipmentTemplate.xlsx"
clear_template_except_header(template_excel)

# 2. 读取原始 Excel
input_file = sys.argv[1]
df = pd.read_excel(input_file)

# 3. 遍历每条订单，创建 ShipItem 对象
ship_items = []
for order_no, group in df.groupby("订单号"):
    first_row = group.iloc[0]  # 代表行（收件人信息用它）

    # 合并 SKU
    merged_sku, need_highlight = merge_skus_with_qty(group)

    full_address = merge_address(
        first_row.get("详细地址1"),
        first_row.get("详细地址2"),
        first_row.get("详细地址3")
    )

    item = ShipItem(
        customer_order=order_no,
        sku=merged_sku,              # 👈 关键变化
        tracking_number="",
        receiver_name=first_row["收货人姓名"],
        receiver_phone=first_row["收货人联系方式"],
        receiver_address=full_address,
        receiver_zip=first_row["收货地址邮编"],
        receiver_city=first_row["城市"],
        receiver_state=first_row["省份"],
        receiver_country=first_row["国家"],
        need_highlight=need_highlight
    )

    ship_items.append(item)

output_dir = create_timestamp_dir("output")

###### 旧模板 #########
output_template = copy_template_with_custom_name(
    template_excel,
    output_dir,
    shipitem_count=len(ship_items)
)

clear_template_except_header(output_template)

fill_template(output_template, 
              ship_items, 
              dict_builder=lambda item: item.to_dict()
)


####### 新模板 #########
# new_output_template = copy_template_with_custom_name(
#     fedex_template_excel,
#     output_dir,
#     shipitem_count=len(ship_items)
# )

# clear_template_except_header(fedex_template_excel)

# fill_template(
#     fedex_template_excel,
#     ship_items,
#     dict_builder=lambda item: item.to_fedex_dict()
# )

print("模板已成功更新并覆盖旧数据")