import time
import openpyxl
import argparse
import sys
from playwright.sync_api import sync_playwright

# 创建 ArgumentParser 对象
parser = argparse.ArgumentParser(description="自动录入题目脚本")

# 添加参数
# 位置参数 1: Excel 文件路径
parser.add_argument(
    "--file_path",
    type=str,
    default=r"D:\电网\新增题库\res_多级层次的存储系统.xlsx",
    help="Excel 题库文件的路径",
)

# 位置参数 2: 目标 URL
parser.add_argument(
    "--target_url",
    type=str,
    default="https://cwgedu.cn/powerTutoring-ui/#/addAnswerSubjectQuestion/index/670/2",
    help="目标网页的 URL",
)

# 解析参数
# 如果用户没有提供参数，argparse 会自动报错并提示用法
args = parser.parse_args()


def run(file_path, target_url):
    print(f"正在读取文件: {file_path} ...")
    try:
        # data_only=True 读取计算后的值
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheet = wb.active
        rows = list(sheet.iter_rows(min_row=2, values_only=True))
        print(f"成功读取 {len(rows)} 条题目数据。")
    except Exception as e:
        print(f"读取 Excel 失败: {e}")
        return

    with sync_playwright() as p:
        print("正在启动浏览器...")
        browser = p.chromium.launch(
            headless=False, slow_mo=500, args=["--start-maximized"]
        )
        context = browser.new_context(no_viewport=True)
        page = context.new_page()

        print(f"正在访问: {target_url}")
        try:
            page.goto(target_url)
        except:
            print("网页加载提示超时，请继续操作...")

        # ================= 手动登录 =================
        print("\n" + "=" * 50)
        print(" 1. 请手动登录。")
        print(" 2. 确保页面停留在列表页（能看到【新增】按钮）。")
        print(" 3. 登录好后，点击本窗口按【回车】开始。")
        print("=" * 50 + "\n")
        input(">>> (按回车键开始...)")
        print(">>> 开始执行...")

        for i, row in enumerate(rows):
            if i < 1:
                continue

            # 辅助函数
            def get_col(index):
                if index < len(row) and row[index] is not None:
                    return str(row[index]).strip()
                return ""

            # Excel 列映射 (根据你之前提供的顺序)
            question_title = get_col(0)
            question_type = get_col(1)
            opt_a = get_col(6)
            opt_b = get_col(8)
            opt_c = get_col(10)
            opt_d = get_col(12)
            answer = get_col(3)
            analysis = get_col(4)

            print(
                f"[{i+1}/{len(rows)}] 录入: {question_title[:10]}... ({question_type})"
            )

            try:
                # 1. 点击“新增”
                page.get_by_text("新增", exact=False).first.click()

                # 2. 锁定“对话框”
                dialog = page.get_by_role("dialog", name="添加答题端题目")
                dialog.wait_for(state="visible", timeout=5000)

                # 3. 填写题目名称
                name_input = dialog.get_by_placeholder("请输入题目名称")
                name_input.fill(question_title)

                # 4. 选择题目类型
                if question_type:
                    # 点击下拉框
                    dialog.get_by_placeholder("请选择").click()
                    # 选择对应的类型
                    page.get_by_role("listitem").filter(
                        has_text=question_type
                    ).first.click()

                    # ========================================================
                    # 【新增逻辑】如果是判断题，点击两次删除按钮
                    # ========================================================
                    if "判断题" in question_type:
                        print("  -> 检测到判断题，正在删除多余选项...")
                        try:
                            # 根据你提供的HTML，删除按钮有 class "remove-link"
                            # 也可以用 text="删除答案选项"
                            remove_btn = dialog.locator(".remove-link")

                            # 点击两次
                            for _ in range(2):
                                remove_btn.click()
                                page.wait_for_timeout(300)  # 稍微等待动画
                        except Exception as e:
                            print(f"  -> 删除选项失败(可能按钮点不到): {e}")

                # 5. 填写选项
                # 判断题通常只有A和B，Excel里C和D应该是空的，fill_option会自动跳过
                def fill_option(letter, content):
                    if content:
                        option_row = dialog.locator(".el-form-item").filter(
                            has_text=f"选项{letter}"
                        )
                        # 检查输入框是否存在(防止判断题删多了或者没删掉导致找不到)
                        if option_row.count() > 0:
                            option_row.get_by_placeholder("请输入内容").fill(content)

                fill_option("A", opt_a)
                fill_option("B", opt_b)
                fill_option("C", opt_c)
                fill_option("D", opt_d)

                # 6. 设置正确答案
                if answer:
                    clean_ans = answer.upper().split(",")
                    for char in clean_ans:
                        try:
                            option_row = dialog.locator(".el-form-item").filter(
                                has_text=f"选项{char}"
                            )
                            option_row.locator("label.el-radio").click()
                        except Exception as e:
                            print(f"  - 设置答案 {char} 失败: {e}")

                # 7. 填写解析
                if analysis:
                    dialog.get_by_placeholder("请输入解析").fill(analysis)

                # 8. 点击“确 定”
                dialog.get_by_role("button", name="确 定").click()

                # 9. 等待弹窗消失
                dialog.wait_for(state="hidden", timeout=5000)
                page.wait_for_timeout(500)

            except Exception as e:
                print(f"!!! 第 {i+1} 条出错: {e}")
                print(">>> 暂停脚本，请检查错误原因 (按 Resume 继续)")
                page.pause()

        print("任务完成。")


if __name__ == "__main__":
    # 调用主函数
    run(args.file_path, args.target_url)
