import os
import pandas as pd

# 設定資料夾路徑
input_folder = "input_excels"
output_folder = "output"


os.makedirs(output_folder, exist_ok=True)

def merge_excels(folder_path):
    all_data = []
    for file in os.listdir(folder_path):
        if file.endswith(".xlsx"):
            df = pd.read_excel(os.path.join(folder_path, file))
            df["來源檔案"] = file
            all_data.append(df)
    return pd.concat(all_data, ignore_index=True)

def generate_monthly_report(df):
    df["日期"] = pd.to_datetime(df["日期"])
    daily_count = df.groupby(df["日期"].dt.date).size().reset_index(name="每日件數")
    monthly_total = daily_count["每日件數"].sum()

    print("本月總件數：", monthly_total)

    return daily_count, monthly_total

# 執行
merged_df = merge_excels(input_folder)
merged_path = os.path.join(output_folder, "merged.xlsx")
merged_df.to_excel(merged_path, index=False)

daily_count_df, total = generate_monthly_report(merged_df)
report_path = os.path.join(output_folder, "monthly_report.xlsx")
daily_count_df.to_excel(report_path, index=False)

print("合併完成!輸出:merged.xlsx 與 monthly_report.xlsx")