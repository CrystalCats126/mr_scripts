"""
描述：
根据给定的excel文件生成分析数据
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# 设置绘图风格
plt.rcParams["font.sans-serif"] = ["SimHei", "Arial Unicode MS"]
plt.rcParams["axes.unicode_minus"] = False
sns.set_style("whitegrid", {"font.sans-serif": ["SimHei", "Arial Unicode MS"]})


def analyze_exam_data_clean_10(file_path):
    # 1. 加载数据
    df = pd.read_csv(file_path)

    # 2. 特征提取
    def extract_info(paper_name):
        company = "其他"
        if any(
            k in paper_name
            for k in [
                "南网",
                "南方电网",
                "广东电网",
                "广西电网",
                "云南电网",
                "贵州电网",
                "海南电网",
            ]
        ):
            company = "南网"
        elif any(k in paper_name for k in ["国网", "国家电网", "江苏三新"]):
            company = "国网"

        subject = "其他"
        if "电工" in paper_name or "电气" in paper_name:
            subject = "电工类"
        elif "通信" in paper_name or "计算机" in paper_name:
            subject = "通信计算机类"
        elif "其他理工" in paper_name:
            subject = "其他理工类"

        return pd.Series([company, subject])

    df[["Company", "Subject"]] = df["试卷名称"].apply(extract_info)

    # 3. 去除噪声 (得分 < 10)
    limit_score = 10
    df_clean = df[df["得分"] >= limit_score].copy()
    removed_count = len(df) - len(df_clean)

    print(f"========== 数据清洗报告 ==========")
    print(f"原始记录: {len(df)}")
    print(f"剔除噪声: {removed_count} 条 (得分 < {limit_score})")
    print(f"有效记录: {len(df_clean)}")

    # 4. 总体统计
    print(f"\n========== 总体情况 (修正后) ==========")
    print(f"平均分: {df_clean['得分'].mean():.2f}")
    print(f"中位数: {df_clean['得分'].median()}")
    print(f"及格率: {(df_clean['得分'] >= 60).mean()*100:.2f}%")

    # 5. 分组对比
    print("\n========== 分组对比 (Top Groups) ==========")
    stats = df_clean.groupby(["Company", "Subject"])["得分"].agg(
        ["count", "mean", "median"]
    )
    print(stats.round(1))

    # 6. 绘图
    plt.figure(figsize=(12, 6))
    plot_data = df_clean[df_clean["Company"].isin(["南网", "国网"])]
    sns.barplot(
        x="Subject",
        y="得分",
        hue="Company",
        data=plot_data,
        estimator="mean",
        errorbar=None,
        palette="viridis",
    )
    plt.title(f"各公司与专业平均成绩对比 (已去除<{limit_score}分噪声)")
    plt.ylabel("平均得分")
    plt.axhline(60, color="red", linestyle="--", alpha=0.5, label="及格线")
    plt.legend()
    plt.tight_layout()
    plt.show()


# 运行调用
if __name__ == "__main__":
    file_name = "records_1764772186233(1).xlsx - 考试记录数据.csv"
    analyze_exam_data_clean_10(file_name)
