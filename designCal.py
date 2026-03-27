import pandas as pd
import numpy as np


def calculate_module_progress(excel_path, save_result_path=None):
    """
    按功能子模块汇总开发进度（不清洗数据，保留原始全量行数）

    参数:
        excel_path: 原始Excel文件路径
        save_result_path: 汇总结果保存路径（可选）

    返回:
        格式化的进度汇总结果列表和汇总结果DataFrame
    """
    # 1. 读取原始Excel数据
    try:
        df = pd.read_excel(excel_path)
        print(f"成功读取原始数据，共 {df.shape[0]} 行 {df.shape[1]} 列")
    except Exception as e:
        print(f"读取Excel文件失败：{str(e)}")
        raise

    # 2. 数据预处理：仅做层级数据填充（不进行任何数据清洗/剔除）
    df_process = df.copy()

    # 2.1 向下填充功能模块和功能子模块（解决层级数据首行标注、后续空值问题）
    # 仅保留填充，不做任何筛选，保留所有原始行
    df_process['功能模块'] = df_process['功能模块'].ffill()
    df_process['功能子模块'] = df_process['功能子模块'].ffill()

    # 【关键修改】取消所有有效行筛选，直接使用填充后的全量数据
    # 不剔除任何行，保持原始总行数不变
    df_full = df_process.reset_index(drop=True)
    print(f"层级数据填充完成，保留全量原始数据共 {df_full.shape[0]} 行（与原始Excel一致）")

    # 3. 处理开发进度数据，识别100%完成（数据中1代表100%完成，0代表未完成）
    def convert_progress(progress_val):
        """转换进度值，统一判断标准"""
        if pd.isna(progress_val):
            return 0
        try:
            progress_num = float(progress_val)
            return 1 if progress_num >= 1 else 0
        except (ValueError, TypeError):
            return 0

    # 新增完成状态列（1=已完成100%，0=未完成），全量数据都参与计算
    df_full['是否完成100%'] = df_full['设计进度'].apply(convert_progress)

    # 4. 按功能子模块分组汇总进度（包含全量数据，包括空值分组）
    progress_summary = []
    module_groups = df_full.groupby(['功能模块', '功能子模块'])

    for (main_module, sub_module), group_data in module_groups:
        # 统计核心数据（基于全量原始行，无剔除）
        total_rows = len(group_data)  # 该子模块原始全量行数
        completed_rows = group_data['是否完成100%'].sum()  # 100%完成行数

        # 计算进度状态（匹配用户要求格式）
        if completed_rows == 0:
            progress_status = "无进度"
        else:
            progress_percent = (completed_rows / total_rows) * 100
            progress_status = f"开发{int(round(progress_percent))}%"

        # 整理结果（包含详细统计和格式化结果）
        progress_summary.append({
            '主模块': main_module,
            '功能子模块': sub_module,
            '总行数（原始全量）': total_rows,
            '100%完成行数': completed_rows,
            '进度状态': progress_status,
            '格式化输出': f"{get_unified_main_module(main_module)}-{sub_module} ：{progress_status}"
        })

    # 5. 数据整理和保存（无encoding参数，保留全量汇总结果）
    summary_df = pd.DataFrame(progress_summary)

    # 保存汇总结果到Excel
    if save_result_path:
        try:
            summary_df.to_excel(save_result_path, index=False)
            print(f"汇总结果已保存至：{save_result_path}")
        except Exception as e:
            print(f"保存Excel文件失败：{str(e)}")
            raise

    # 6. 输出用户要求格式的进度结果
    print("\n" + "=" * 80)
    print("按功能子模块汇总进度结果（基于全量原始数据）")
    print("=" * 80)

    # 按用户示例顺序排序输出
    for item in sorted(progress_summary, key=lambda x: (x['主模块'].find('数字住建'), x['功能子模块'])):
        # 过滤掉子模块为NaN的无效分组（可选，如需保留可删除该判断）
        if pd.notna(item['功能子模块']):
            print(item['格式化输出'])

    return progress_summary, summary_df


def get_unified_main_module(main_module):
    """统一主模块名称，匹配用户要求的输出格式"""
    if pd.isna(main_module):
        return "未知模块"
    elif '数字住建' in main_module:
        return "数字住建统一门户"
    elif '业务支撑' in main_module:
        return "业务支撑系统"
    else:
        return main_module


# ---------------------- 调用示例 ----------------------
if __name__ == "__main__":
    # 替换为你的实际Excel文件路径
    INPUT_EXCEL_PATH = r"D:\\DESKTOP\\交付功能点细化.xlsx"  # 替换为实际文件路径
    OUTPUT_EXCEL_PATH = r"D:\\DESKTOP\\功能子模块进度汇总_最终可运行版.xlsx"  # 输出结果路径

    # 执行进度计算
    try:
        progress_result, summary_dataframe = calculate_module_progress(
            excel_path=INPUT_EXCEL_PATH,
            save_result_path=OUTPUT_EXCEL_PATH
        )
        print("\n进度计算完成，结果已返回并保存！")
    except Exception as e:
        print(f"运行出错：{str(e)}")
        print("请检查Excel文件路径是否正确，以及文件格式是否符合要求。")