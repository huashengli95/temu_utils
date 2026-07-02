import pandas as pd
import sys
import random
import re
import os
import shutil
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.cell.cell import MergedCell
from pathlib import Path
import json

# Load settings from ./config
STATE_MAP = {
    "NL": "NF",
    "QC": "PQ",
}

# 渠道代码和业务线路配置
CHANNEL_CODES = ["1252", "1253", "0106", "6064"]
BUSINESS_ROUTES = ["CA", "CA-PURO", "AUS"]
DEFAULT_BUSINESS_ROUTE = "CA"
DEFAULT_CARRIER_TYPE = 1  # 承运商/模版类型
DEFAULT_SUB_TEMPLATE_TYPE = 1  # 子模版类型

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

def build_output_filename(shipitem_count, type, ext=".xlsx"):
    """
    生成: YYYYMMDD_<count>.xlsx
    """
    date_str = datetime.now().strftime("%Y%m%d")
    return f"{date_str}_{shipitem_count}_{type}{ext}"

def copy_template_with_custom_name(template_path, target_dir, shipitem_count, type):
    """
    拷贝模板到目录，并按规则重命名
    """
    new_filename = build_output_filename(shipitem_count, type)
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

        # 4. 仅由 '-' 组���的占位符（-, --, --- 等）
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
            cell = ws.cell(row=row_idx, column=col_idx)
            
            # 检查是否是合并单元格，如果是则跳过
            if not isinstance(cell, MergedCell):
                cell.value = item_dict.get(header, "")
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
                 need_highlight=False, channel_code=None, business_route=None,
                 sku_qty=None):
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
        
        # 新模板字段
        self.channel_code = channel_code or random.choice(CHANNEL_CODES)
        self.business_route = business_route or DEFAULT_BUSINESS_ROUTE
        self.sku_qty = sku_qty or 1  # 数量

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
    
    def to_xiaolong_dict(self):
        """
        生成 xiaolong_template.xlsx 格式的数据
        """
        return {
            "订单号(必填)": self.customer_order,
            "渠道代码(必填)": self.channel_code,
            "业务线路(必填:CA/CA-PURO/CA-LIGHT等)": self.business_route,
            "承运商/模版类型(CA填1-4; PURO填1-2; CA-LIGHT填1=CA,2=PURO,3=FEDEX)": 1,
            "子模版类型(仅CA-LIGHT选1或2时填,代表具体模板1-4或1-2)": 1,
            "SKU(必填)": self.sku,
            "数量(必填)": self.sku_qty,
            "包裹长cm(必填)": self.package_spec.length,
            "包裹宽cm(必填)": self.package_spec.width,
            "包裹高cm(必填)": self.package_spec.height,
            "重量kg(必填)": self.package_spec.weight,
            "发件人-姓名(必填)": self.sender["name"],
            "发件人-电话(必填)": self.sender["phone"],
            "发件人-国家(必填)": self.sender["country"],
            "发件人-地址1(必填)": self.sender["address"],
            "发件人-地址2(选填)": "",
            "发件人-城市(必填)": self.sender["city"],
            "发件人-省份(必填)": normalize_state_for_new_template(self.sender["state"]),
            "发件人-邮编(必填)": self.sender["zip"],
            "收件人-姓名(必填)": self.receiver_name,
            "收件人-电话(必填)": clean_temu_phone(self.receiver_phone),
            "收件人-国家(地址用,必填)": "CA",
            "收件人-地址1(必填)": self.receiver_address,
            "收件人-地址2(选填)": "",
            "收件人-城市(必填)": self.receiver_city,
            "收件人-省份(必填)": normalize_state_for_new_template(self.receiver_state),
            "收件人-邮编(必填)": self.receiver_zip,
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

# ###### 旧模板 #########
# to_agent_output_template = copy_template_with_custom_name(
#     template_excel,
#     output_dir,
#     shipitem_count=len(ship_items),
#     type="to_agent"
# )

# clear_template_except_header(to_agent_output_template)

# fill_template(to_agent_output_template, 
#               ship_items, 
#               dict_builder=lambda item: item.to_dict()
# )


# ####### 新模板 #########
# to_fedex_output_template = copy_template_with_custom_name(
#     fedex_template_excel,
#     output_dir,
#     shipitem_count=len(ship_items),
#     type="to_fedex"
# )

# clear_template_except_header(to_fedex_output_template)

# fill_template(
#     to_fedex_output_template,
#     ship_items,
#     dict_builder=lambda item: item.to_fedex_dict()
# )

####### xiaolong 新模板 #########
xiaolong_template_excel = BASE_DIR / "xiaolong_template.xlsx"
to_xiaolong_output_template = copy_template_with_custom_name(
    xiaolong_template_excel,
    output_dir,
    shipitem_count=len(ship_items),
    type="to_xiaolong"
)

clear_template_except_header(to_xiaolong_output_template)

fill_template(
    to_xiaolong_output_template,
    ship_items,
    dict_builder=lambda item: item.to_xiaolong_dict()
)

print("模板已成功更新并覆盖旧数据")
print(f"已生成 {len(ship_items)} 条订单数据到以下模板:")
# print(f"  - {to_agent_output_template}")
# print(f"  - {to_fedex_output_template}")
print(f"  - {to_xiaolong_output_template}")
