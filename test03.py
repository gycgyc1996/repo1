import pandas as pd

# 读取Excel文件的两个Sheet
file_path = 'Z:\\NAS02\\GYC共享\\Python\\Excel测试\\import.xlsx'
df_sales = pd.read_excel(file_path, sheet_name='Sheet1')  # 假设销售数据在第一个工作表
df_regions = pd.read_excel(file_path, sheet_name='Sheet2')  # 假设地区信息在第二个工作表

# 打印读取的数据以检查
print("销售数据:")
print(df_sales)
print("\n地区人员:")
print(df_regions)

# 合并两个DataFrame，基于地区连接
df_merged = pd.merge(df_sales, df_regions, on='地区', how='left')

# 输出合并后的数据，检查合并是否正确
print("\n合并后的数据:")
print(df_merged)

# 将处理结果输出到新的Excel文件
output_file_path = 'Z:\\NAS02\\GYC共享\\Python\\Excel测试\\export.xlsx'
df_merged.to_excel(output_file_path, index=False)
print(f"\n结果已导出到 {output_file_path}")
