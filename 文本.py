import pandas as pd
import tkinter as tk
from tkinter import filedialog

# 弹出文件选择框
def select_file(title):
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xlsx *.xls")])
    return file_path

print("请选择【空表格】（目标表）")
target_file = select_file("选择空表格")

print("请选择【数据来源表格】")
source_file = select_file("选择数据来源表格")

# 读取表格
target_df = pd.read_excel(target_file, header=0)
source_df = pd.read_excel(source_file)

# 👉 这里设置数据源的列名（根据你的实际修改）
row_key_col = source_df.columns[0]   # 行字段
col_key_col = source_df.columns[1]   # 列字段
value_col   = source_df.columns[2]   # 值

# 转成字典，加快匹配速度
data_dict = {}
for _, row in source_df.iterrows():
    key = (row[row_key_col], row[col_key_col])
    data_dict[key] = row[value_col]

# 获取行列标签
row_labels = target_df.iloc[:, 0]     # 第一列（行名）
col_labels = target_df.columns[1:]    # 第一行（列名）

# 填充数据
for i, row_name in enumerate(row_labels):
    for col_name in col_labels:
        key = (row_name, col_name)
        if key in data_dict:
            target_df.loc[i, col_name] = data_dict[key]

# 保存结果（覆盖原文件）
target_df.to_excel(target_file, index=False)

print("✅ 填充完成，已保存到原文件！")