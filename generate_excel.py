# Create an Excel scoring from input data and package it into a ZIP for reliable download

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import zipfile
import os

# Define paths
base_dir = os.path.dirname(os.path.abspath(__file__))
input_excel_path = os.path.join(base_dir, "Mike-prepared.xlsx")
output_excel_path = os.path.join(base_dir, "telecom_sales_scoring_model.xlsx")
zip_path = os.path.join(base_dir, "telecom_sales_scoring_model.zip")

# Load input data
wb_in = openpyxl.load_workbook(input_excel_path, data_only=True)
ws_in = wb_in.active

# Find column indices
headers_in = [cell.value for cell in ws_in[1]]
col_idx = {name: idx for idx, name in enumerate(headers_in)}

# Create output workbook
wb_out = Workbook()
ws_out = wb_out.active
ws_out.title = "Sales_Data"

headers_out = [
    "区域",
    "2024收入",
    "2025收入",
    "收入增长率",
    "客户数",
    "区域企业数",
    "客户经理数",
    "人均产值",
    "客户渗透率",
    "规模得分",
    "增长得分",
    "市场得分",
    "效率得分",
    "战略得分",
    "潜力得分",
    "综合得分",
    "排名"
]

# Write headers
for col, h in enumerate(headers_out, start=1):
    cell = ws_out.cell(row=1, column=col, value=h)
    cell.font = Font(bold=True)

# Write data and add formulas
data_rows = list(ws_in.iter_rows(min_row=2, values_only=True))
last_row = len(data_rows) + 1
if last_row < 2:
    last_row = 2

for r, row_data in enumerate(data_rows, start=2):
    # Extract needed values based on columns
    region = row_data[col_idx['区县']] if '区县' in col_idx else ""
    rev_2024 = row_data[col_idx['区域内24年收入']] if '区域内24年收入' in col_idx else 0
    rev_2025 = row_data[col_idx['2025年工业收入（万元）']] if '2025年工业收入（万元）' in col_idx else 0
    clients = row_data[col_idx['客户数']] if '客户数' in col_idx else 0
    non_clients = row_data[col_idx['区域内未成为客户的数量']] if '区域内未成为客户的数量' in col_idx else 0
    managers = row_data[col_idx['客户经理数']] if '客户经理数' in col_idx else 0
    
    total_enterprises = (clients or 0) + (non_clients or 0)
    
    # Write values
    ws_out[f"A{r}"] = region
    ws_out[f"B{r}"] = rev_2024
    ws_out[f"C{r}"] = rev_2025
    ws_out[f"E{r}"] = clients
    ws_out[f"F{r}"] = total_enterprises
    ws_out[f"G{r}"] = managers

    # Add formulas
    ws_out[f"D{r}"] = f"=IF(B{r}=0,0,(C{r}-B{r})/B{r})"           # 收入增长率
    ws_out[f"H{r}"] = f"=IF(G{r}=0,0,C{r}/G{r})"                   # 人均产值
    ws_out[f"I{r}"] = f"=IF(F{r}=0,0,E{r}/F{r})"                   # 客户渗透率
    
    ws_out[f"J{r}"] = f"=IF(MAX($C$2:$C${last_row})=0,0,C{r}/MAX($C$2:$C${last_row})*25)"                  # 规模得分
    ws_out[f"K{r}"] = f"=IF(MAX($D$2:$D${last_row})=0,0,D{r}/MAX($D$2:$D${last_row})*20)"                  # 增长得分
    ws_out[f"L{r}"] = f"=IF(MAX($I$2:$I${last_row})=0,0,I{r}/MAX($I$2:$I${last_row})*15)"                  # 市场得分
    ws_out[f"M{r}"] = f"=IF(MAX($H$2:$H${last_row})=0,0,H{r}/MAX($H$2:$H${last_row})*15)"                  # 效率得分
    
    ws_out[f"N{r}"] = 0                                           # 战略得分 (手动填写)
    ws_out[f"O{r}"] = 0                                           # 潜力得分 (手动填写)
    
    ws_out[f"P{r}"] = f"=SUM(J{r}:O{r})"                           # 综合得分
    ws_out[f"Q{r}"] = f"=RANK(P{r},$P$2:$P${last_row},0)"                  # 排名

# Adjust column widths
for i in range(1, len(headers_out)+1):
    ws_out.column_dimensions[get_column_letter(i)].width = 16

# Save Excel
wb_out.save(output_excel_path)

# Zip it
with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as z:
    z.write(output_excel_path, os.path.basename(output_excel_path))

print(f"Generated {output_excel_path} and {zip_path}")
