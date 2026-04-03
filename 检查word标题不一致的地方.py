import win32com.client as win32
import os
import shutil


def clear_com_cache():
    """清理 win32com 缓存，防止接口识别错误"""
    import sys
    gen_py_path = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Temp', 'gen_py')
    if os.path.exists(gen_py_path):
        try:
            shutil.rmtree(gen_py_path)
        except:
            pass


def get_headings_safe(file_path):
    headings = []
    abs_path = os.path.abspath(file_path)

    # 确保文件存在
    if not os.path.exists(abs_path):
        print(f"找不到文件: {abs_path}")
        return []

    try:
        # 使用更通用的方式获取 Word 对象
        word = win32.DispatchEx('Word.Application')
        word.Visible = False
        word.DisplayAlerts = 0

        # 只传路径和只读参数
        doc = word.Documents.Open(abs_path, False, True)

        for p in doc.Paragraphs:
            style_name = p.Style.NameLocal
            text = p.Range.Text.strip().replace('\r', '').replace('\n', '')

            # 判断逻辑：包含“标题”或“Heading”字样，或者确实带自动编号
            number = p.Range.ListFormat.ListString
            is_title = "标题" in style_name or "Heading" in style_name

            if (is_title or (number and len(number) > 0)) and text:
                headings.append(f"{number} {text}".strip())

        doc.Close(False)
    except Exception as e:
        print(f"读取 {os.path.basename(file_path)} 失败: {e}")
    return headings


def run_compare():
    # 1. 先尝试清理缓存（解决 Open() 参数报错）
    clear_com_cache()

    path_a = r"D:\DESKTOP\生成追加_仅留标题.docx"
    path_b = r"D:\DESKTOP\模板 目录 全部_仅留标题.docx"

    print("--- 启动比对程序 ---")
    list1 = get_headings_safe(path_a)
    list2 = get_headings_safe(path_b)

    if not list1 and not list2:
        print("❌ 错误：未能从任何文档中提取到标题。")
        print("提示：请确认文档中的标题是否真的应用了“标题”样式。")
        return

    print(f"文档A 标题数: {len(list1)}")
    print(f"文档B 标题数: {len(list2)}")
    print("-" * 30)

    max_len = max(len(list1), len(list2))
    diff_count = 0
    for i in range(max_len):
        t1 = list1[i] if i < len(list1) else "【缺失】"
        t2 = list2[i] if i < len(list2) else "【缺失】"

        if t1 != t2:
            diff_count += 1
            print(f"差异 {diff_count} (位置 {i + 1}):")
            print(f"  A: {t1}")
            print(f"  B: {t2}")
            print("-" * 20)

    if diff_count == 0:
        print("✅ 结果：标题和编号完全匹配！")
    else:
        print(f"比对结束：共发现 {diff_count} 处不同。")


if __name__ == "__main__":
    try:
        run_compare()
    finally:
        # 尝试清理残留进程
        try:
            word = win32.Dispatch('Word.Application')
            word.Quit()
        except:
            pass