"""
DeepSeek 智能选题分类脚本 (增强通信类识别版)
修改点：
1. 优化 Prompt：增加了【领域判断指南】，强制模型先区分是通信还是计算机。
2. 关键词引导：明确告知模型哪些关键词属于通信（如SDH、光纤、调制、基站等）。
3. 调试输出：处理时会打印前几条的分类结果，方便你观察是否识别正确。
"""

import pandas as pd
import requests
import os
import re
import time
import argparse
import glob
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm

# ================= 配置区域 =================

# 1. API Key
API_KEY = ""

# 2. 题目列名
QUESTION_COLUMN = "题目名称"

# 3. 线程数 (建议 10)
MAX_WORKERS = 15

# 4. 最大重试次数
MAX_RETRIES = 5

# 5. 输出文件夹
OUTPUT_DIR = "南网_分类结果_按详细知识点拆分"
OUTPUT_MASTER_FILE = "题库_总表(含分类结果).xlsx"

# ===========================================

API_URL = "https://api.deepseek.com/chat/completions"

# === 核心修改：增强型 Prompt ===
CATEGORIES_PROMPT = """
你是一个精通【计算机科学】和【通信工程】的双领域专家。
请对给出的考试题目进行分类。

!!! 必须严格遵循以下判断逻辑 !!!
Step 1: 先判断题目属于【通信】还是【计算机】领域：
   - 【通信类】特征：涉及信号处理、调制解调(ASK/FSK/PSK)、频谱、信道、香农公式、光纤(SDH/OTN/PON)、移动通信(5G/4G/LTE/GSM/基站)、卫星、微波、交换技术(电路交换)、运营商网络设备等。
   - 【计算机类】特征：涉及数据结构(链表/树/栈)、算法、编程语言(Java/C/Python)、数据库(SQL/Oracle)、操作系统(进程/死锁)、软件工程(敏捷/测试)、计算机体系结构(流水线/Cache)等。
   - 【模糊地带】：如果涉及"网络"，请区分：
       -> 偏物理层、传输设备、光缆、电信网架构 -> 归入【通信类】
       -> 偏TCP/IP协议栈、Web应用、HTTP/DNS -> 归入【计算机类】

Step 2: 根据 Step 1 的判断，从下方对应的列表中选择最精确的一个分类。

Step 3: 仅返回“编号+类别名称”（例如：“5 调制解调技术” 或 “2 线性表”），不要解释。

====== 分类列表 ======

【A. 通信类列表 (优先匹配)】
[通信原理信号与系统]
1 信号与系统的基本概念
2 信号与系统的时域及频域特性分析
3 通信与通信系统的基本概念
4 信道特性及复用、多址、均衡、分集技术
5 调制解调技术
6 数字信号的最佳接收
7 信源编解码技术
8 信道编解码技术
9 通信系统同步技术
10 通信标准及组织
[光纤通信]
11 光纤通信技术基础
12 光纤传输技术（SDH、WDM、OTN、fgOTN、PTN）
13 光纤的结构与特性、光缆结构
14 光纤通信系统常用器件
15 光纤通信系统收发架构
16 接入技术（光纤接入 PON、无线接入）
17 光性能测量与监控仪器仪表
[数据通信网-通信侧]
18 数据通信网络体系架构
19 交换技术（电路交换、分组交换、ATM）
20 常用数据通信网络协议
21 数据通信网络设备/接口
22 网络安全（通信网安全模型、加密）
23 数据网新技术基础（MPLS，IPv6，SDN）
24 数据网组网通用配置分析
[移动通信及其他]
25 移动通信系统基础
26 5G/6G 移动通信关键技术基础
27 卫星通信基础及应用
28 基于 H.320、H.323、SIP 协议的会议电视系统
29 电力线载波通信
30 新一代电力应急指挥通信系统
31 物联网技术及应用
32 通信网络智能管理

【B. 计算机类列表】
[数据结构与算法]
1 数据结构基本概念
2 线性表
3 栈和队列
4 数组、数组的压缩存储与字符串
5 树和二叉树
6 图
7 查找
8 内部排序
9 算法设计与分析基础
[数据库系统]
10 数据库基本概念
11 关系数据库基本理论
12 关系数据库标准语言 SQL
13 事务处理和并发控制
14 备份和恢复
15 数据库应用系统设计与开发
[计算机网络-IT侧]
16 计算机网络体系结构
17 物理层
18 数据链路层
19 网络层
20 传输层
21 应用层
22 网络管理与网络安全
23 无线网络与移动网络
[操作系统]
24 系统运行环境和运行机制
25 进程与线程管理
26 内存管理
27 文件管理
28 设备管理
29 操作系统安全与保护
[计算机组成与体系结构]
30 计算机系统概述
31 数据的机器级表示和运算
32 多级层次的存储系统
33 指令系统
34 中央处理器
35 总线与输入输出系统
36 并行计算架构
[软件工程]
37 软件工程基本概念
38 软件开发过程管理
39 常见软件开发方法
40 需求分析
41 系统设计
42 系统开发
43 软件测试
44 软件交付与维护
[信息新技术]
45 人工智能基础
46 物联网基础
47 大数据基础
"""

parser = argparse.ArgumentParser(description="混合题库智能分类脚本")
parser.add_argument(
    "-i",
    "--input",
    type=str,
    default=r"E:\my_script\专业知识 南方电网通信计算机类题库",
    help="输入的文件路径 或 文件夹路径",
)
config = parser.parse_args()


def get_all_excel_files(path):
    file_list = []
    if os.path.isfile(path):
        file_list.append(path)
    elif os.path.isdir(path):
        print(f"正在扫描文件夹: {path} ...")
        for ext in ["*.xlsx", "*.xls", "*.csv"]:
            file_list.extend(glob.glob(os.path.join(path, "**", ext), recursive=True))
        file_list = list(set(file_list))
    else:
        print(f"❌ 路径不存在: {path}")
    return [f for f in file_list if not os.path.basename(f).startswith("~$")]


def call_deepseek_api(question_text, index_info=""):
    if not question_text or str(question_text).strip() == "":
        return "未分类"

    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {API_KEY}"}

    messages = [
        {
            "role": "system",
            "content": "你是一个严谨的考试题目分类专家。请仔细区分通信工程与计算机科学的边界。",
        },
        {
            "role": "user",
            "content": f"{CATEGORIES_PROMPT}\n\n题目内容：{question_text}\n\n所属分类（仅输出编号和名称）：",
        },
    ]

    data = {
        "model": "deepseek-chat",
        "messages": messages,
        "temperature": 0.1,
        "max_tokens": 60,
    }

    last_error = ""
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = requests.post(API_URL, headers=headers, json=data, timeout=30)
            if response.status_code == 200:
                content = response.json()["choices"][0]["message"]["content"].strip()
                content = (
                    content.replace("【A】", "").replace("【B】", "").strip()
                )  # 清洗可能出现的前缀
                return content
            elif response.status_code == 429:
                time.sleep(2**attempt)
                continue
            else:
                last_error = f"HTTP {response.status_code}"
        except Exception as e:
            last_error = f"网络异常: {str(e)}"

        if attempt < MAX_RETRIES:
            time.sleep(1.5 * attempt)

    return f"分类失败 [{last_error}]"


def process_single_task(index, question):
    return index, call_deepseek_api(question, index_info=index)


def load_and_merge_data(file_list):
    all_dfs = []
    print(f"找到 {len(file_list)} 个文件，开始读取...")
    for f in file_list:
        try:
            if f.endswith(".csv"):
                try:
                    df = pd.read_csv(f, encoding="utf-8")
                except:
                    df = pd.read_csv(f, encoding="gbk")
            else:
                df = pd.read_excel(f)

            df["来源文件"] = os.path.basename(f)
            if QUESTION_COLUMN in df.columns:
                all_dfs.append(df)
            else:
                print(
                    f"⚠️ 跳过文件 (未找到'{QUESTION_COLUMN}'列): {os.path.basename(f)}"
                )
        except Exception as e:
            print(f"❌ 读取失败 {os.path.basename(f)}: {e}")

    if not all_dfs:
        return pd.DataFrame()
    return pd.concat(all_dfs, ignore_index=True)


def split_excel_by_category(df):
    print(f"\n正在拆分结果到文件夹: {OUTPUT_DIR} ...")
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    df_clean = df[
        df["知识点分类"].notna()
        & (df["知识点分类"] != "")
        & (~df["知识点分类"].astype(str).str.contains("失败|错误|未分类"))
    ]

    if df_clean.empty:
        print("没有有效数据，无法拆分。")
        return

    grouped = df_clean.groupby("知识点分类")
    count = 0
    for name, group in grouped:
        safe_name = str(name).strip()
        safe_name = re.sub(r"[\n\r\t]", "_", safe_name)
        safe_name = re.sub(r'[\\/*?:"<>|]', "_", safe_name)
        safe_name = safe_name[:60]  # 限制长度

        file_path = os.path.join(OUTPUT_DIR, f"{safe_name}.xlsx")
        try:
            group.to_excel(file_path, index=False)
            count += 1
        except Exception as e:
            print(f"⚠️ 保存 '{safe_name}' 失败: {e}")

    print(f"拆分完成！共 {count} 个分类文件。")


def main():
    input_path = config.input
    files = get_all_excel_files(input_path)
    if not files:
        return

    # 优先读取中间结果
    if os.path.exists(OUTPUT_MASTER_FILE):
        print(f"读取中间文件 '{OUTPUT_MASTER_FILE}' 继续处理...")
        df = pd.read_excel(OUTPUT_MASTER_FILE)
    else:
        df = load_and_merge_data(files)

    if df.empty:
        return

    if "知识点分类" not in df.columns:
        df["知识点分类"] = ""

    # 强制重跑失败或未分类的项
    unprocessed_mask = (
        (df["知识点分类"].isna())
        | (df["知识点分类"] == "")
        | (df["知识点分类"].astype(str).str.contains("失败|错误"))
    )

    unprocessed_indices = df[unprocessed_mask].index.tolist()
    total_tasks = len(unprocessed_indices)
    print(f"待处理: {total_tasks} 条")

    if total_tasks > 0:
        print(f"开始并发分类 (线程:{MAX_WORKERS})...")
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            future_to_idx = {
                executor.submit(
                    process_single_task, idx, df.loc[idx, QUESTION_COLUMN]
                ): idx
                for idx in unprocessed_indices
            }

            pbar = tqdm(
                as_completed(future_to_idx), total=total_tasks, unit="题", ncols=100
            )
            counter = 0
            for future in pbar:
                idx, result = future.result()
                df.at[idx, "知识点分类"] = result
                counter += 1

                # 打印前几条结果，方便用户Debug
                if counter <= 5:
                    tqdm.write(
                        f"Debug: 题目片段 '{str(df.loc[idx, QUESTION_COLUMN])[:10]}...' -> 分类: {result}"
                    )

                if "失败" in result:
                    pbar.set_postfix_str(f"ID:{idx} 失败")
                if counter % 50 == 0:
                    df.to_excel(OUTPUT_MASTER_FILE, index=False)

        df.to_excel(OUTPUT_MASTER_FILE, index=False)
        print("\n分类完成。")

    split_excel_by_category(df)
    print(f"\n全部完成！结果在: {os.path.abspath(OUTPUT_DIR)}")


if __name__ == "__main__":
    main()
