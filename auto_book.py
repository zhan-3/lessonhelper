from playwright.sync_api import sync_playwright
import pandas as pd
import os


def get_chrome_path():
    paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for p in paths:
        if os.path.exists(p):
            return p
    return None


def auto_book_all_labs():
    chrome_path = get_chrome_path()
    results = []
    success_count = 0
    fail_count = 0

    with sync_playwright() as p:
        if chrome_path:
            browser = p.chromium.launch(headless=False, executable_path=chrome_path)
        else:
            browser = p.chromium.launch(headless=False)

        context = browser.new_context(viewport={"width": 1920, "height": 1080})
        page = context.new_page()

        page.goto("http://openlab.hitwh.edu.cn/dxwl/booking/#/booking")
        page.wait_for_timeout(5000)

        page.locator("button >> text=选课").click()
        page.wait_for_timeout(2000)

        exp_cells = page.locator("table tr td:first-child").all()
        exp_names = []
        for cell in exp_cells:
            text = cell.text_content().strip()
            if text and "20" in text:
                exp_names.append(text)
        exp_names = list(set(exp_names))[:15]
        print(f"共找到 {len(exp_names)} 个实验")

        for exp_name in exp_names:
            print(f"\n处理实验：{exp_name}")
            try:
                page.locator(f"td:has-text('{exp_name}')").first.click()
                page.wait_for_timeout(1500)

                date_cells = page.locator("table tr td:first-child").all()
                dates = []
                for cell in date_cells:
                    text = cell.text_content().strip()
                    if text.startswith("2026-"):
                        dates.append(text)

                for date in dates:
                    try:
                        page.locator(f"td:has-text('{date}')").first.click()
                        page.wait_for_timeout(1500)

                        sections = page.locator("text=/大节/").all()
                        for section in sections:
                            try:
                                section.click()
                                page.wait_for_timeout(500)
                                page.locator("button >> text=查询").click()
                                page.wait_for_timeout(1500)

                                dialog = page.locator("dialog")
                                if dialog.is_visible():
                                    dialog_text = dialog.text_content()
                                    if "没有可供选择" in dialog_text:
                                        print(
                                            f"  {date} {section.text_content().strip()} 已满"
                                        )
                                        fail_count += 1
                                    else:
                                        print(f"预约成功：{exp_name} {date}")
                                        success_count += 1
                                    page.locator("button >> text=确认").click()
                                    page.wait_for_timeout(500)

                                results.append(
                                    {
                                        "实验名称": exp_name,
                                        "日期": date,
                                        "节次": section.text_content().strip(),
                                        "状态": "预约成功" if success_count else "已满",
                                    }
                                )

                            except Exception as e:
                                print(f"节次处理失败：{str(e)}")
                                fail_count += 1

                        page.locator("button >> text=刷新").click()
                        page.wait_for_timeout(1000)

                    except Exception as e:
                        print(f"日期处理失败：{date}")
                        page.reload()
                        page.wait_for_timeout(2000)

                page.locator("button >> text=返回").first.click()
                page.wait_for_timeout(1000)

            except Exception as e:
                print(f"实验处理失败：{exp_name}")
                page.reload()
                page.wait_for_timeout(2000)

        df = pd.DataFrame(results)
        df.to_excel("lab_booking_results.xlsx", index=False)
        print(f"\n完成！成功 {success_count} 个，失败 {fail_count} 个")
        print("结果已保存到：lab_booking_results.xlsx")

        browser.close()


if __name__ == "__main__":
    auto_book_all_labs()
