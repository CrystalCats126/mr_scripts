"""
通过抓包的方式将网络包里面的内容生成excel文件。

（
从指定文件(data.txt)当中的 JSON 数据，并生成 Excel 题目文件。
）当中读取文件
操作步骤：

在页面按 F12 打开开发者工具。

点击顶部的 “Network” (网络) 标签。

选中下方的 “Fetch/XHR” 过滤器(这一步是为了过滤掉图片、CSS等杂乱信息,只看数据)。

刷新页面(按 F5)。
点击那个文件，在右侧看 “Preview” (预览)。

观察左侧出现的文件列表。寻找paperDetail.json 文件。


如果你看到类似  question_list 这样的结构，这就是题库源数据。

导出：把 question_list 里的内容全部复制出来,执行该脚本即可提取出文件。
找出这个文件中你认为答案明确不正确的题目或者解析题目对不上的题目,不用在意答案格式错误, 最终结果以表格形式给出（需要给出题目内容和修改建议）
查看一下这个文件里面的题目解析有没有问题，把有问题的以表格形式给出，需要给出题目内容，仅查看解析。
"""

import json
import pandas as pd

# 1. 读取文件（注意加上 encoding='utf-8' 防止报错）
# 请确保 data.txt 文件在同一目录下
try:
    with open("mr_scripts\data.txt", "r", encoding="utf-8") as f:
        json_data_str = f.read()
except FileNotFoundError:
    print("错误：找不到 data.txt 文件，请确认文件名和路径。")
    exit()


def parse_questions_to_excel(json_str, output_filename="题库导出_已过滤.xlsx"):
    try:
        # 解析JSON数据
        data = json.loads(json_str)

        rows = []

        # 定义题目类型映射
        type_mapping = {1: "单选题", 2: "多选题", 3: "判断题"}

        for item in data:
            # 获取原始类型ID
            raw_type = item.get("type")

            # 获取类型名称，如果不在映射表中，返回 "未知类型"
            q_type = type_mapping.get(raw_type, "未知类型")

            # 【关键修改】在这里判断：如果是未知类型，直接跳过进入下一轮循环
            if q_type == "未知类型":
                continue

            row = {}

            # 1. 基础信息提取
            row["题目类型"] = q_type
            row["题目内容"] = item.get("name", "").strip()
            row["正确答案"] = item.get("rightAnswer", "").strip()
            row["题目解析"] = item.get("analysis", "").strip()

            # 2. 选项提取
            option_keys = ["选项A", "选项B", "选项C", "选项D", "选项E", "选项F"]
            for key in option_keys:
                row[key] = ""

            options = item.get("questionOptionList", [])
            if options:
                for opt in options:
                    tag = opt.get("optionTag", "")
                    content = opt.get("optionContent", "")
                    if tag in row:
                        row[tag] = content
                    else:
                        if tag.upper() in ["A", "B", "C", "D", "E", "F"]:
                            row[f"选项{tag.upper()}"] = content

            rows.append(row)

        if not rows:
            print("没有提取到有效题目，请检查JSON数据或类型映射。")
            return

        # 转换为 DataFrame
        df = pd.DataFrame(rows)

        # 排列列顺序
        columns_order = [
            "题目类型",
            "题目内容",
            "正确答案",
            "解析",
            "选项A",
            "选项B",
            "选项C",
            "选项D",
        ]
        final_cols = [col for col in columns_order if col in df.columns] + [
            col for col in df.columns if col not in columns_order
        ]
        df = df[final_cols]

        # 导出到 Excel
        df.to_excel(output_filename, index=False)
        print(f"成功导出文件：{output_filename}")
        print(f"共提取了 {len(rows)} 道题目（已自动过滤未知类型）。")

    except json.JSONDecodeError as e:
        print(f"JSON解析错误: {e}")
    except Exception as e:
        print(f"发生错误: {e}")


if __name__ == "__main__":
    parse_questions_to_excel(json_data_str)
