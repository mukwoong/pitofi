import pandas as pd
import os

# ——————————————————数据部分————————————————————————————————————
source_folder = os.path.join(os.getcwd(), 'PI')
xlsx_files = [f for f in os.listdir(source_folder) if f.endswith('.xlsx')]
# 读取文件
file_path = os.path.join(source_folder, xlsx_files[0])
df = pd.read_excel(file_path)


# 颜色结束行
null_values = df.iloc[10:, 3].isnull()
color_end_row = null_values.idxmax() if null_values.any() else None

# 有效区间[7行：颜色结束行，所有列]
subset = df.iloc[10:color_end_row, :]

# 工厂
subset.iloc[:, -1] = subset.iloc[:, -1].ffill()
# 款号
subset.iloc[:, 0] = subset.iloc[:, 0] .ffill()

# 品牌名字
filename = xlsx_files[0]
prefix = filename.split('-')[0]

new_order = [-3, 0, -1, -4, -2]

subset = subset.iloc[:, new_order]

subset.insert(2, 'des', value=None)
subset.insert(3, 'com', value=None)
subset.insert(6, 'ship', value=None)
subset.insert(7, 'brand', value=prefix)

subset.to_excel('output.xlsx')