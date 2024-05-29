import pandas as pd
import numpy as np

# 读取平衡面板Excel文件
df_path = 'C:/Users/14470/Desktop/保险balance.xlsx'
df = pd.read_excel(df_path)

# 解析保险代码以获取年份、地级市代码和保险公司名称
df['年份'] = df['保险代码'].str[:4].astype(int)
df['地级市代码'] = df['保险代码'].str[4:10]
df['保险公司'] = df['保险代码'].str[10:]

df.sort_values(by=['地级市代码', '保险公司', '年份'], inplace=True)

# 计算保费和赔付差异比例，也即term1
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

    # 填充NaN值（如果需要）
    df.fillna(0, inplace=True)
    return df

# 应用函数到DataFrame
df = calculate_previous_year_proportions(df, ['保费', '赔付'])

# 定义函数以计算除了当前地级市以外的总和的对数差分,也即term2
def calculate_log_diff(df, column):
    # 计算每个保险公司在所有地级市每年的总和
    total_sum = df.groupby(['保险公司', '年份'])[column].transform('sum')
    # 计算每个保险公司在当前地级市每年的总和
    city_sum = df.groupby(['保险公司', '地级市代码', '年份'])[column].transform('sum')
    # 在除了当前地级市以外的总和
    df[f'{column}_其他地级市总和'] = total_sum - city_sum
    # 对这个总和进行对数变换，加1以避免取对数0的情况
    df[f'{column}_对数'] = np.log10(df[f'{column}_其他地级市总和']+0.01)
    # 计算对数值的年度差分
    df[f'{column}_对数差分'] = df.groupby(['保险公司', '地级市代码'])[f'{column}_对数'].diff().fillna(0)
    # 修正比例中的无穷大或未定义值
    df.replace([np.inf, -np.inf], np.nan, inplace=True)
    df.fillna(0, inplace=True)
    return df

# 分别对保费和赔付进行处理
df = calculate_log_diff(df, '保费')
df = calculate_log_diff(df, '赔付')



# 计算每个地级市的B值
df['保费B值'] = df['保费比例'] * df['保费_对数差分']
df['赔付B值'] = df['赔付比例'] * df['赔付_对数差分']

# 汇总每个地级市每年的B值
summary = df.groupby(['地级市代码', '年份']).agg({'保费B值': 'sum', '赔付B值': 'sum'}).reset_index()

# 保存到Excel文件
output_file_path = 'C:/Users/14470/Desktop/保险data.xlsx'
summary.to_excel(output_file_path, index=False)