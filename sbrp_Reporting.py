import os
from glob import glob
import pandas as pd
from datetime import datetime
import shutil

# 1. 设置工作默认路径
default_path = "/Users/linyunxia/Documents/report"
os.chdir(default_path)

# 2. 获取当前时间和保存路径
current_time = datetime.now()
save_base = os.path.join(
    default_path,
    "result",
    f"{current_time.year}",
    f"{current_time.month}",
    f"{current_time.day}",
    f"{current_time.hour}"
)
os.makedirs(save_base, exist_ok=True)  # 创建保存路径（如果不存在）

# 3. 移动tmp文件夹中的所有文件到保存路径
tmp_folder = "tmp"
if os.path.exists(tmp_folder):
    for file in glob(os.path.join(tmp_folder, "*")):
        shutil.move(file, save_base)

# 4. 读取并合并tmp文件夹中的Excel文件（此时文件已移动到save_base）
all_files = glob(os.path.join(save_base, "*.xlsx"))
dfs = []
for file in all_files:
    df = pd.read_excel(file)
    dfs.append(df)

merged_df = pd.concat(dfs, ignore_index=True)

# 5. 读取基准Excel文件并提取需要的列
org_folder = "org"
base_file = os.path.join(default_path, org_folder, "GC Hotel System List 2024.xlsx")
base_df = pd.read_excel(base_file, sheet_name="Details")

selected_columns = base_df[["Inncode", "ITM", "IT E-mail1", "IT E-mail2", "RPA RMH"]]

# 6. 根据Inncode进行匹配
final_df = pd.merge(merged_df, selected_columns, on="Inncode", how="left")

# 7. 保存中间文件和最终报告到save_base路径
temp_file = os.path.join(save_base, "temp.xlsx")
final_report = os.path.join(save_base, "final_report.xlsx")

merged_df.to_excel(temp_file, index=False)
final_df.to_excel(final_report, index=False)

print("文件已成功处理并保存至:", save_base)