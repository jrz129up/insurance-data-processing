import pandas as pd
import numpy as np

df_path = 'xlsx'
df = pd.read_excel(df_path)

df['年份'] = df['保险代码'].str[:4].astype(int)
df['地级市代码'] = df['保险代码'].str[4:10]
df['保险公司'] = df['保险代码'].str[10:]

df.sort_values(by=['地级市代码', '保险公司', '年份'], inplace=True)

# compute term1
def calculate_previous_year_proportions(df, columns):
    for column in columns:
        # 计算上一年每个地级市每个保险公司的指定列（保费或赔付）总额
        df[f'上一年{column}'] = df.groupby(['地级市代码', '保险公司'])[column].shift(1)
        # 计算每个地级市每年的指定列的总和
        df['年度总和'] = df.groupby(['地级市代码', '年份'])[column].transform('sum')
        # 将每年的总和赋给下一年作为“上一年的总和”
        df[f'上一年地级市{column}总和'] = df.groupby(['地级市代码','保险公司'])['年度总和'].shift(1)
        # 计算除了当前保险公司以外的指定列总和
        df[f'上一年除本公司外{column}总和'] = df[f'上一年地级市{column}总和'] - df[f'上一年{column}']
        # 计算指定列的比例
        df[f'{column}比例'] = df[f'上一年{column}'] / df[f'上一年除本公司外{column}总和']
        # 处理可能出现的分母为0的情况
        df[f'{column}比例'] = df[f'{column}比例'].replace([np.inf, -np.inf], np.nan)

    df.fillna(0, inplace=True)
    return df

# apply the function in order to create a new DataFrame
df = calculate_previous_year_proportions(df, ['保费', '赔付'])

# compute term2
def calculate_log_diff(df, column):
    total_sum = df.groupby(['保险公司', '年份'])[column].transform('sum')
    city_sum = df.groupby(['保险公司', '地级市代码', '年份'])[column].transform('sum')
    df[f'{column}_其他地级市总和'] = total_sum - city_sum
    df[f'{column}_对数'] = np.log10(df[f'{column}_其他地级市总和']+0.01)
    df[f'{column}_对数差分'] = df.groupby(['保险公司', '地级市代码'])[f'{column}_对数'].diff().fillna(0)
    # 修正比例中的无穷大或未定义值
    df.replace([np.inf, -np.inf], np.nan, inplace=True)
    df.fillna(0, inplace=True)
    return df

# deal with 保费 and 赔付 respectively
df = calculate_log_diff(df, '保费')
df = calculate_log_diff(df, '赔付')



# compute each B value
df['保费B值'] = df['保费比例'] * df['保费_对数差分']
df['赔付B值'] = df['赔付比例'] * df['赔付_对数差分']

# accumulate
summary = df.groupby(['地级市代码', '年份']).agg({'保费B值': 'sum', '赔付B值': 'sum'}).reset_index()

# save
output_file_path = 'xlsx'
summary.to_excel(output_file_path, index=False)
