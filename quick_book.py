from playwright.sync_api import sync_playwright
import os

chrome_paths = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]


def quick_book():
    chrome_path = None
    for p in chrome_paths:
        if os.path.exists(p):
            chrome_path = p
            break

    with sync_playwright() as p:
        if chrome_path:
            browser = p.chromium.launch(headless=False, executable_path=chrome_path)
        else:
            browser = p.chromium.launch(headless=False)

        context = browser.new_context(viewport={"width": 1920, "height": 1080})
        page = context.new_page()

        page.goto("http://openlab.hitwh.edu.cn/dxwl/booking/#/booking")
        page.wait_for_timeout(3000)

        print("点击选课...")
        page.locator("button:has-text('选课')").click()
        page.wait_for_timeout(2000)

        print("查找DIY电磁混合磁悬浮...")
        diy_cell = page.locator("td:has-text('DIY电磁混合磁悬浮')").first
        if diy_cell.count() > 0:
            diy_cell.click()
            page.wait_for_timeout(2000)
            print("已选中")

            for target_date in ["2026-04-10", "2026-05-08", "2026-05-15"]:
                print(f"尝试 {target_date}...")
                date_cell = page.locator(f"td:has-text('{target_date}')").first
                if date_cell.count() > 0:
                    date_cell.click()
                    page.wait_for_timeout(2000)
                    page.locator("button:has-text('查 询')").click()
                    page.wait_for_timeout(2000)

                    dialog = page.locator(".ant-modal-content")
                    if dialog.is_visible():
                        content = dialog.text_content()
                        if (
                            "没有可供选择" in content
                            or "已满" in content
                            or "人数已满" in content
                        ):
                            print(f"{target_date} 已满")
                            page.locator("button:has-text('确定')").click()
                            page.wait_for_timeout(1000)
                        else:
                            print(f"{target_date} 可预约!")
                            seat = page.locator("input[type='radio']").first
                            if seat.count() > 0:
                                seat.click()
                                page.wait_for_timeout(500)
                                page.locator("button:has-text('确认')").click()
                                page.wait_for_timeout(3000)
                                print("预约完成!")
                                browser.close()
                                return
                    else:
                        page.wait_for_timeout(3000)

            print("DIY满了，尝试其他课程...")
            page.locator("button:has-text('返回')").click()
            page.wait_for_timeout(2000)

            for course in ["磁阻效应", "表面张力", "偏振光"]:
                print(f"尝试 {course}...")
                course_cell = page.locator(f"td:has-text('{course}')").first
                if course_cell.count() > 0:
                    course_cell.click()
                    page.wait_for_timeout(2000)

                    for target_date in ["2026-04-10", "2026-05-08", "2026-04-24"]:
                        date_cell = page.locator(f"td:has-text('{target_date}')").first
                        if date_cell.count() > 0:
                            date_cell.click()
                            page.wait_for_timeout(2000)
                            page.locator("button:has-text('查 询')").click()
                            page.wait_for_timeout(2000)

                            dialog = page.locator(".ant-modal-content")
                            if dialog.is_visible():
                                content = dialog.text_content()
                                if (
                                    "没有可供选择" not in content
                                    and "已满" not in content
                                    and "人数已满" not in content
                                ):
                                    print(f"{course} {target_date} 可预约!")
                                    seat = page.locator("input[type='radio']").first
                                    if seat.count() > 0:
                                        seat.click()
                                        page.wait_for_timeout(500)
                                        page.locator("button:has-text('确认')").click()
                                        page.wait_for_timeout(3000)
                                        print("预约完成!")
                                        browser.close()
                                        return
                                else:
                                    page.locator("button:has-text('确定')").click()
                                    page.wait_for_timeout(1000)

                page.locator("button:has-text('返回')").click()
                page.wait_for_timeout(2000)

        page.screenshot(path="booking_state.png")
        print("截图已保存: booking_state.png")
        browser.close()


if __name__ == "__main__":
    quick_book()
