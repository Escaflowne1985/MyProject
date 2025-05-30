import pandas as pd
import os

# 读取 Excel 文件
file_path = "menu.xlsx"  # 替换为你的文件路径
df = pd.read_excel(file_path)

# 仅保留必要字段，并剔除空值
df_clean = df[['文章专栏', '专栏分类', '组合标题', '序号']].dropna(subset=['文章专栏', '专栏分类', '组合标题', '序号'])

# 确保序号为整数类型并排序
df_clean['序号'] = df_clean['序号'].astype(int)
df_clean = df_clean.sort_values(by='序号')

# 获取每个专栏的最小序号，作为命名编号
first_index_map = df_clean.groupby('文章专栏')['序号'].min().to_dict()

# 构造带编号的专栏名映射（用于统一排序和命名）
column_items = sorted(first_index_map.items(), key=lambda x: x[1])
column_index_map = {col: f"{i+1:02d}" for i, (col, _) in enumerate(column_items)}

# 输出目录
output_dir = "articles_by_column"
os.makedirs(output_dir, exist_ok=True)

# 遍历专栏生成 md 文件
for column_name, group_df in df_clean.groupby('文章专栏'):
    index_prefix = column_index_map[column_name]
    filename = f"{index_prefix}.{column_name}.md"
    filepath = os.path.join(output_dir, filename)

    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(f"# {column_name}\n\n")
        
        # 按专栏分类输出表格
        for category, cat_group in group_df.groupby('专栏分类'):
            f.write(f"## {category}\n\n")
            f.write("| 编号 | 内容和链接 |\n")
            f.write("| ---- | ----------- |\n")
            for row in cat_group.itertuples():
                f.write(f"| {row.序号:04d} | {row.组合标题.strip()} |\n")
            f.write("\n")  # 分类之间空一行

            
            
import pandas as pd
import os
import re
from urllib.parse import quote  # ✅ 用于路径中的空格和特殊字符转义

# 读取原始表格，获取所有专栏
df = pd.read_excel("menu.xlsx")
columns_all = df['文章专栏'].dropna().unique()

# 扫描输出目录中所有合规命名的 Markdown 文件
output_dir = "articles_by_column"
pattern = re.compile(r'^(\d{2})\.(.+)\.md$')
valid_files = {}

for file in os.listdir(output_dir):
    match = pattern.match(file)
    if match:
        index, column_title = match.groups()
        safe_path = quote(f"./{output_dir}/{file}")  # ✅ 路径转义处理
        valid_files[column_title.strip()] = (index, safe_path)

# 生成 README 表格
index_md = "# 专栏汇总索引\n\n"
index_md += "| 编号 | 专栏名称 | 阅读链接 |\n"
index_md += "| ---- | -------- | -------- |\n"

for i, col in enumerate(columns_all, start=1):
    col = col.strip()
    if col in valid_files:
        index, path = valid_files[col]
        index_md += f"| {index} | {col} | [查看]({path}) |\n"
    else:
        index_md += f"| {i:02d} | {col} | 查看 |\n"

# 写入 README.md
readme_path = os.path.join(output_dir, "README.md")
with open(readme_path, 'w', encoding='utf-8') as f:
    f.write(index_md)

print(f"✅ README 生成完毕：{readme_path}")