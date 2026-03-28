import pandas as pd
import os


def rename_duplicates_with_numbers(excel_path, output_path):
    # 1. 加载 Excel
    # header=0 表示第一行是表头，如果你的表头在第2行，请改为 header=1
    df = pd.read_excel(excel_path, sheet_name='Sheet1')

    # 2. 清洗列名：去除首尾空格、换行符，防止出现找不到列名的情况
    df.columns = [str(c).strip() for c in df.columns]
    print(f"检测到的有效列名: {list(df.columns)}")

    # 检查必要的列是否存在
    required_cols = ['系统', '功能模块', '功能点']
    for col in required_cols:
        if col not in df.columns:
            # 如果还是找不到，尝试模糊匹配（比如只要包含“功能模块”四个字就行）
            matched = [c for c in df.columns if col in c]
            if matched:
                df.rename(columns={matched[0]: col}, inplace=True)
                print(f"已将列 '{matched[0]}' 自动修正为 '{col}'")
            else:
                raise KeyError(f"在Excel中找不到包含 '{col}' 的列，请检查表头名称。")

    # 3. 处理合并单元格产生的 NaN
    df['系统'] = df['系统'].ffill()
    df['功能模块'] = df['功能模块'].ffill()

    # 注意：功能点如果是逐行填写的，不建议用 ffill()，否则会把空白行也填上重复内容
    # 但根据你的文件结构，如果功能点也有合并，请取消下面这行的注释
    # df['功能点'] = df['功能点'].ffill()

    # 4. 定义智能编号函数
    def apply_numbering(series):
        counts = {}
        new_values = []
        # 先统计总数，确定哪些需要编号
        total_counts = series.value_counts()

        for val in series:
            val_str = str(val).strip()
            if val_str == 'nan' or not val_str or val_str == 'None':
                new_values.append(val)  # 保持原样（空值）
                continue

            # 只有出现次数 > 1 的才编号
            if total_counts[val] > 1:
                counts[val_str] = counts.get(val_str, 0) + 1
                new_values.append(f"{val_str}{counts[val_str]}")
            else:
                new_values.append(val_str)
        return new_values

    # 5. 执行编号逻辑
    print("正在处理功能模块重复项...")
    df['功能模块'] = apply_numbering(df['功能模块'])

    print("正在处理功能点重复项...")
    df['功能点'] = apply_numbering(df['功能点'])

    # 6. 保存
    df.to_excel(output_path, index=False)
    print(f"成功！已处理重复项并保存至: {output_path}")


  # dev :苗:2026-03-28 21:12:04

# --- 路径配置 ---
input_file = r"D:\DESKTOP\功能清单分工.xlsx"
output_file = r"D:\DESKTOP\功能清单分工_已编号.xlsx"

if __name__ == "__main__":
    try:
        rename_duplicates_with_numbers(input_file, output_file)
    except Exception as e:
        print(f"运行出错: {e}")