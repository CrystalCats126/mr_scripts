# -*- coding: utf-8 -*-

"""
Excel 同名文件合并工具
--------------------------------------------------
功能：比较两个文件夹，找出文件名相同的 Excel 文件，
      将它们的内容垂直合并（追加），并保存到指定输出目录。

用法示例：
    python merge_excel_cli.py -a "path/to/folder_A" -b "path/to/folder_B" -o "path/to/output"
"""

import os
import argparse
import pandas as pd
from pathlib import Path
from typing import Set, List


def parse_args():
    """
    解析命令行参数
    """
    parser = argparse.ArgumentParser(
        description="将两个文件夹中同名的 Excel 文件进行合并。",
        formatter_class=argparse.RawTextHelpFormatter,
    )

    parser.add_argument(
        "-a",
        "--dir_a",
        type=str,
        default="E:\my_script\南网_分类结果_按详细知识点拆分",
        help="输入文件夹 A 的路径 (基础数据)",
    )

    parser.add_argument(
        "-b",
        "--dir_b",
        type=str,
        default="E:\my_script\国网通信_分类结果_按详细知识点拆分",
        help="输入文件夹 B 的路径 (要追加的数据)",
    )

    parser.add_argument(
        "-o",
        "--output",
        type=str,
        default="E:\my_script\国网通信_final_version",
        help="输出结果文件夹的路径",
    )

    return parser.parse_args()


def get_excel_filenames(folder_path: Path) -> Set[str]:
    """
    获取指定文件夹下所有有效 Excel 文件名的集合。
    过滤掉以 ~$ 开头的临时文件。
    """
    if not folder_path.exists():
        print(f"❌ 错误：文件夹不存在 -> {folder_path}")
        return set()

    # 匹配 .xls 和 .xlsx，忽略大小写
    files = []
    extensions = ["*.xlsx", "*.xls", "*.XLSX", "*.XLS"]

    for ext in extensions:
        # 使用 pathlib 的 glob 查找
        found = list(folder_path.glob(ext))
        files.extend(found)

    # 提取文件名，并过滤掉临时文件
    filenames = {f.name for f in files if not f.name.startswith("~$")}
    return filenames


def merge_files(dir_a: Path, dir_b: Path, output_dir: Path):
    """
    执行文件合并的核心逻辑
    """
    # 1. 获取文件列表
    print(f"正在扫描文件夹 A: {dir_a}")
    files_a = get_excel_filenames(dir_a)

    print(f"正在扫描文件夹 B: {dir_b}")
    files_b = get_excel_filenames(dir_b)

    # 2. 取交集（找出两边都有的文件）
    common_files = files_a.intersection(files_b)

    if not common_files:
        print("\n⚠️  未在两个文件夹中找到同名的 Excel 文件，程序结束。")
        return

    print(f"\n✅ 找到 {len(common_files)} 个同名文件，准备合并...")

    # 3. 确保输出目录存在
    output_dir.mkdir(parents=True, exist_ok=True)

    # 4. 遍历处理
    success_count = 0
    fail_count = 0

    for idx, filename in enumerate(common_files, 1):
        file_path_a = dir_a / filename
        file_path_b = dir_b / filename
        output_path = output_dir / filename

        print(f"[{idx}/{len(common_files)}] 处理: {filename}")

        try:
            # 读取数据
            # 使用 keep_default_na=True 保持默认空值处理
            df_a = pd.read_excel(file_path_a)
            df_b = pd.read_excel(file_path_b)

            # 记录原始行数
            len_a = len(df_a)
            len_b = len(df_b)

            # 合并数据 (A 在上，B 在下)
            # sort=False 防止列名顺序改变
            merged_df = pd.concat([df_a, df_b], ignore_index=True, sort=False)

            # 保存数据
            merged_df.to_excel(output_path, index=False)

            print(
                f"   -> 合并成功: A({len_a}行) + B({len_b}行) = 总计({len(merged_df)}行)"
            )
            success_count += 1

        except Exception as e:
            print(f"   -> ❌ 失败: {str(e)}")
            fail_count += 1

    # 5. 总结
    print("\n" + "=" * 30)
    print(f"处理完成！")
    print(f"成功: {success_count}")
    print(f"失败: {fail_count}")
    print(f"结果保存在: {output_dir.resolve()}")
    print("=" * 30)


def main():
    # 1. 解析参数
    args = parse_args()

    # 2. 转换为 Path 对象，方便操作
    path_a = Path(args.dir_a)
    path_b = Path(args.dir_b)
    path_out = Path(args.output)

    # 3. 执行合并
    merge_files(path_a, path_b, path_out)


if __name__ == "__main__":
    main()
