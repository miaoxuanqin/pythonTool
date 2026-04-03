import re
import os


def get_pure_text(line):
    """
    使用正则表达式移除行首的编号部分。
    例如: "2.4.1.1. 系统概述" -> "系统概述"
    """
    line = line.strip()
    if not line:
        return ""
    # ^[\d\.]+ 匹配开头所有的数字和点
    # \s* 匹配编号后的空格
    pure_text = re.sub(r'^[\d\.]+\s*', '', line)
    return pure_text


def find_first_diff_ignore_numbers(path1, path2):
    if not os.path.exists(path1) or not os.path.exists(path2):
        print("错误：请确认文件路径是否正确。")
        return

    print(f"正在比对文字内容（已忽略编号差异）...")

    try:
        # 使用 utf-8 编码，并忽略无法识别的字符
        with open(path1, 'r', encoding='utf-8', errors='ignore') as f1, \
                open(path2, 'r', encoding='utf-8', errors='ignore') as f2:

            line_num = 0
            while True:
                line_num += 1
                l1 = f1.readline()
                l2 = f2.readline()

                # 检查是否同时到达文件末尾
                if not l1 and not l2:
                    print("✅ 比对完成：两个文件的文字内容完全一致。")
                    break

                # 检查文件长度是否一致
                if not l1 or not l2:
                    print(f"❌ 发现不一致！位置：第 {line_num} 行")
                    print(f"原因：文件 {'txt2.txt' if not l1 else 'txt.txt'} 提前结束。")
                    break

                # 提取纯文字部分进行比对
                text1 = get_pure_text(l1)
                text2 = get_pure_text(l2)

                if text1 != text2:
                    print(f"❌ 发现文字不一致！起始行号：第 {line_num} 行")
                    print(f"--- 详细对比 ---")
                    print(f"txt.txt  (A) 原始行: {l1.strip()}")
                    print(f"txt.txt  (A) 提取文字: {text1}")
                    print(f"\ntxt2.txt (B) 原始行: {l2.strip()}")
                    print(f"txt2.txt (B) 提取文字: {text2}")
                    print(f"----------------")
                    break  # 找到第一处不同即停止

    except Exception as e:
        print(f"程序运行出错: {e}")


# 配置路径
file_a = r"D:\DESKTOP\txt.txt"
file_b = r"D:\DESKTOP\txt2.txt"

if __name__ == "__main__":
    find_first_diff_ignore_numbers(file_a, file_b)