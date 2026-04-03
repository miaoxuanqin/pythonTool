import win32com.client as win32
import os


def read_titles_with_numbers(file_path):
    # 确保路径是绝对路径，win32com 对相对路径支持不好
    abs_path = os.path.abspath(file_path)

    # 启动 Word 应用程序
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False  # 不显示 Word 界面

    try:
        doc = word.Documents.Open(abs_path)
        print(f"成功读取文档：{os.path.basename(file_path)}\n")

        for para in doc.Paragraphs:
            # 判断段落是否有编号 或 是否为标题样式
            # ListString 会抓取自动编号生成的字符串（如 "1.1"）
            number_prefix = para.Range.ListFormat.ListString
            text = para.Range.Text.strip().replace('\r', '').replace('\n', '')

            # 我们只打印有编号的内容，或者特定标题样式的段落
            if number_prefix or "Heading" in para.Style.NameLocal or "标题" in para.Style.NameLocal:
                # 拼接编号和文本
                full_title = f"{number_prefix} {text}".strip()
                if full_title:
                    print(full_title)

        doc.Close()
    except Exception as e:
        print(f"读取失败: {e}")
    finally:
        word.Quit()


# 执行读取
path = r'D:\DESKTOP\模板 目录 全部_仅留标题.docx'
read_titles_with_numbers(path)