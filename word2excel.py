"""
将 Word 文档转换为 Excel 格式
"""

import pandas as pd
from docx import Document
import re
import os
import argparse

parser = argparse.ArgumentParser(description="将word文档转换为Excel格式")
parser.add_argument(
    "--input_file", type=str, default="26一批本科真题.docx", help="Word 文件路径"
)
config = parser.parse_args()


def parse_docx(file_path):
    """
    解析 Word 文档，提取题目、答案和选项
    """
    document = Document(file_path)
    questions = []
    current_q = None

    # 选项正则：以 A-F 开头，后面跟点、顿号或空格
    option_start_pattern = re.compile(r"^\s*([A-F])\s*[.、\s]\s*")

    for para in document.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # 1. 检查是否为选项行
        is_option_line = False
        if current_q and option_start_pattern.match(text):
            is_option_line = True

        if is_option_line and current_q:
            # 处理一行可能有多个选项的情况 (如: A xxx B xxx)
            markers = ["A", "B", "C", "D", "E", "F"]
            found_markers = []
            for m in markers:
                # 查找所有选项标记的位置 (必须在行首或前面有空格)
                iter_pattern = r"(?:^|\s+)(" + m + r")\s*[.、\s]"
                for match in re.finditer(iter_pattern, text):
                    found_markers.append((match.start(), match.group(1)))

            found_markers.sort()

            for i in range(len(found_markers)):
                start_idx, label = found_markers[i]
                # 结束位置是下一个标记的开始，或者是行尾
                if i < len(found_markers) - 1:
                    end_idx = found_markers[i + 1][0]
                else:
                    end_idx = len(text)

                # 提取内容并清洗
                sub_text = text[start_idx:end_idx]
                # 去除选项标签前缀 (如 "A ")
                clean_content = re.sub(
                    r"^\s*" + label + r"\s*[.、\s]\s*", "", sub_text
                ).strip()

                current_q["options"][label] = clean_content

        else:
            # 2. 检查是否为新题目
            # 特征：包含括号（全角或半角），且行尾有大写字母（答案）
            has_open = "（" in text or "(" in text
            has_close = "）" in text or ")" in text

            if has_open and has_close:
                # 查找行尾的答案 (如 A, B, ACD)
                ans_match = re.search(r"([A-F]+)\s*$", text)
                if ans_match:
                    ans = ans_match.group(1)
                    # 简单校验：答案必须由 A-F 组成
                    if all(c in "ABCDEF" for c in ans):
                        # 保存上一题
                        if current_q:
                            questions.append(current_q)

                        # 提取题目文本（去掉末尾答案）
                        q_text = text[: ans_match.start()].strip()

                        # 判断类型
                        q_type = ".多选题" if len(ans) > 1 else ".单选题"

                        current_q = {
                            "text": q_text,
                            "type": q_type,
                            "answer": ans,
                            "options": {},
                        }

    # 添加最后一题
    if current_q:
        questions.append(current_q)

    return questions


def convert_to_df(questions):
    """
    将解析的数据转换为目标 DataFrame 格式
    """
    rows = []
    for q in questions:
        # 尝试检测判断题 (如果 A 是正确/True, B 是错误/False)
        opt_a = q["options"].get("A", "")
        opt_b = q["options"].get("B", "")
        if ("正确" in opt_a or "True" in opt_a) and (
            "错误" in opt_b or "False" in opt_b
        ):
            if q["type"] == ".单选题":
                q["type"] = ".判断题"

        row = {
            "题目名称": q["text"],
            "题目类型": q["type"],
            "图片": "",
            "正确答案": q["answer"],
            "解析": "",
            "解析图片": "",
        }

        # 填充选项
        labels = ["A", "B", "C", "D", "E", "F"]
        for i, label in enumerate(labels):
            idx = i + 1
            row[f"选项标签{idx}"] = label if label in q["options"] else ""
            row[f"选项内容{idx}"] = q["options"].get(label, "")

        rows.append(row)

    # 确保列名与模板一致
    cols = [
        "题目名称",
        "题目类型",
        "图片",
        "正确答案",
        "解析",
        "解析图片",
        "选项标签1",
        "选项内容1",
        "选项标签2",
        "选项内容2",
        "选项标签3",
        "选项内容3",
        "选项标签4",
        "选项内容4",
        "选项标签5",
        "选项内容5",
        "选项标签6",
        "选项内容6",
    ]

    df = pd.DataFrame(rows)

    # 补全缺失的列
    for c in cols:
        if c not in df.columns:
            df[c] = ""

    return df[cols]


if __name__ == "__main__":
    input_file = config.input_file
    file_name, ext = os.path.splitext(input_file)
    output_file = file_name + ".xlsx"

    if os.path.exists(input_file):
        print(f"正在处理文件: {input_file} ...")
        questions = parse_docx(input_file)
        print(f"共解析出 {len(questions)} 道题目。")

        df = convert_to_df(questions)
        df.to_excel(output_file, index=False)
        print(f"转换成功！文件已保存为: {output_file}")
    else:
        print(f"错误: 找不到文件 {input_file}，请确保文件在当前目录下。")
