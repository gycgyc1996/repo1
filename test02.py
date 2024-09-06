import pandas as pd

# 读取Excel文件
file_path = 'Z:\\NAS02\\GYC共享\\Python\\Excel测试\\import.xlsx'  # 替换为您的Excel文件路径
df = pd.read_excel(file_path)

# 查看前几行数据，确保读取正确
print("前几行数据：")
print(df.head())

# 统计销售总量
total_sales = df['销售量'].sum()
print(f"销售总量：{total_sales}")

# 计算每月增长率
df['上月销售量'] = df['销售量'].shift(1)
df['增长率'] = ((df['销售量'] - df['上月销售量']) / df['上月销售量']) * 100

# 将统计结果导出到新的Excel文件
output_file_path = 'Z:\\NAS02\\GYC共享\\Python\\Excel测试\\export.xlsx'
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Sales Data', index=False)
    summary_df = pd.DataFrame({
        '指标': ['销售总量'],
        '数值': [total_sales]
    })
    summary_df.to_excel(writer, sheet_name='Summary', index=False)

print(f"分析结果已导出到 {output_file_path}")
