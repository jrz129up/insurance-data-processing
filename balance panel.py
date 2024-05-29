import pandas as pd
from itertools import product

# 加载Excel文件的Sheet2
df = pd.read_excel('C:/Users/14470/Desktop/12-21年市保险公司数据.xlsx')

# 提取年份、地级市代码和保险公司名称
df['年份'] = df['保险代码'].str.slice(0, 4).astype(int)
df['地级市代码'] = df['保险代码'].str.slice(4, 10)
df['保险公司'] = df['保险代码'].str.slice(10)

# 获取唯一的年份、地级市代码和保险公司名称
years = df['年份'].unique()
cities = df['地级市代码'].unique()
companies = df['保险公司'].unique()

# 创建所有可能的年份、地级市和保险公司组合
all_combinations = pd.DataFrame(list(product(years, cities, companies)), columns=['年份', '地级市代码', '保险公司'])

# 将原始数据与所有可能的组合进行合并，缺失的组合将保费和赔付设置为0
balanced_panel = all_combinations.merge(df, on=['年份', '地级市代码', '保险公司'], how='left')
balanced_panel['保费'] = balanced_panel['保费'].fillna(0)
balanced_panel['赔付'] = balanced_panel['赔付'].fillna(0)

# 重建保险代码列
balanced_panel['保险代码'] = (balanced_panel['年份'].astype(str) + balanced_panel['地级市代码'] + balanced_panel['保险公司'])

# 选择需要的列
balanced_panel = balanced_panel[['保险代码', '保费', '赔付']]

# 保存处理后的数据到新的Excel文件中
balanced_panel.to_excel('C:/Users/14470/Desktop/保险balance.xlsx', index=False)

