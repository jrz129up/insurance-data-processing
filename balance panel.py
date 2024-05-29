import pandas as pd
from itertools import product

df = pd.read_excel('xlsx')

df['年份'] = df['保险代码'].str.slice(0, 4).astype(int)
df['地级市代码'] = df['保险代码'].str.slice(4, 10)
df['保险公司'] = df['保险代码'].str.slice(10)
years = df['年份'].unique()
cities = df['地级市代码'].unique()
companies = df['保险公司'].unique()

all_combinations = pd.DataFrame(list(product(years, cities, companies)), columns=['年份', '地级市代码', '保险公司'])

balanced_panel = all_combinations.merge(df, on=['年份', '地级市代码', '保险公司'], how='left')
balanced_panel['保费'] = balanced_panel['保费'].fillna(0)
balanced_panel['赔付'] = balanced_panel['赔付'].fillna(0)

balanced_panel['保险代码'] = (balanced_panel['年份'].astype(str) + balanced_panel['地级市代码'] + balanced_panel['保险公司'])

balanced_panel = balanced_panel[['保险代码', '保费', '赔付']]

balanced_panel.to_excel('xlsx', index=False)

