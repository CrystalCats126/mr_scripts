import time
from playwright.sync_api import sync_playwright


def run():
    # ================= 配置区域 =================
    TARGET_URL = "https://cwgedu.cn/powerTutoring-ui/#/answer/correction"
    # ===========================================

    with sync_playwright() as p:
        print("正在启动浏览器...")
        browser = p.chromium.launch(headless=False, slow_mo=500)
        context = browser.new_context()
        page = context.new_page()

        print(f"正在访问: {TARGET_URL}")
        page.goto(TARGET_URL)

        # === 手动登录等待区 ===
        print("\n" + "=" * 50)
        print(" 1. 请手动登录。")
        print(" 2. 确保位于第一页。")
        print(" 3. 准备好后，回到底部终端按【回车键】开始。")
        print("=" * 50 + "\n")
        input(">>> (按回车键开始...)")

        # 当前所在的页码，默认从第1页开始
        current_page_num = 1

        while True:
            print(f"\n>>> [第 {current_page_num} 页] 正在检查数据...")

            # --- [内层循环]：清理当前页面的数据 ---
            while True:
                # 1. 查找包含“该类目已删除”的行
                # 使用 .first 锁定第一条，防止元素引用失效
                target_row = page.locator("tr").filter(has_text="该类目已删除").first

                if target_row.count() == 0:
                    print(f"--- 第 {current_page_num} 页清理完毕 ---")
                    break

                print(f"发现删除项，正在处理...")
                try:
                    target_row.get_by_text("删除").last.click()

                    # 处理二次确认弹窗
                    try:
                        confirm_btn = page.get_by_role("button", name="确定")
                        if confirm_btn.is_visible(timeout=2000):
                            confirm_btn.click()
                        else:
                            page.get_by_role("button", name="确认").click(timeout=1000)
                    except:
                        pass

                    # 稍微等待表格刷新
                    page.wait_for_timeout(1000)

                except Exception as e:
                    print(f"操作重试中: {e}")
                    page.wait_for_timeout(1000)

            # --- [外层循环]：数字翻页逻辑 ---

            # 目标是下一页的数字
            next_page_num = current_page_num + 1
            print(f"正在寻找第 {next_page_num} 页的按钮...")

            # ============================================================
            # 【关键点：精准定位数字按钮】
            # 我们需要找到文本完全等于 "2", "3" 等数字的元素。
            # 为了防止误点表格里的数字，这里尝试限定在分页容器内查找。
            # ============================================================

            # 策略：尝试查找文本完全匹配 next_page_num 的元素
            # exact=True 表示精确匹配，不会因为搜 "2" 而匹配到 "20"
            next_btn = page.get_by_text(str(next_page_num), exact=True)

            # 进阶定位（推荐）：如果你发现它老是点错，请取消下面这行的注释，并根据F12修改 class
            # next_btn = page.locator(".ant-pagination, .el-pagination, .pagination").get_by_text(str(next_page_num), exact=True)

            if next_btn.count() > 0 and next_btn.is_visible():
                # 有时候数字存在但可能是不可点的文本（比如 "共 2 页"），所以最好检查它是不是 role=listitem 或 button
                # 这里直接尝试点击，如果报错说明不是按钮
                try:
                    next_btn.click()
                    current_page_num += 1
                    print(f">>> 成功点击第 {current_page_num} 页，等待加载...")

                    # 翻页后必须等待，给页面一点加载时间
                    page.wait_for_timeout(2000)
                except Exception as e:
                    print(
                        f">>> 找到了数字 {next_page_num} 但无法点击，可能已到末尾或被遮挡。"
                    )
                    print(f"错误信息: {e}")
                    break
            else:
                # 找不到下一个数字了
                print(
                    f">>> 页面上没有找到数字 '{next_page_num}'，推测已到达最后一页。任务结束。"
                )
                break

        print("所有页面处理完毕。")


if __name__ == "__main__":
    run()
