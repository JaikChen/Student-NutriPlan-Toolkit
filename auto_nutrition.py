import time
import os
import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# ================= 配置区域 =================
# 获取当前脚本所在目录
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
# 自动定位到 manager_inventory.py 生成结果的目录
FOLDER_PATH = os.path.join(CURRENT_DIR, 'data', '2_食材入库管理', '输出结果')

# 目标网址
TARGET_URL = "https://yyjh.xszz.edu.cn/yygsjh/dlsp/cgqdwhSchool"


# ===========================================

def get_academic_info(date_str):
    """根据日期判断学年和学期"""
    try:
        date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        year = date_obj.year
        month = date_obj.month

        # 逻辑：2-8月是春季（属于上一年的学年），9-1月是秋季
        if 2 <= month <= 8:
            return f"{year - 1}-{year}", "春季学期"
        elif month >= 9:
            return f"{year}-{year + 1}", "秋季学期"
        else:  # month == 1
            return f"{year - 1}-{year}", "秋季学期"
    except Exception as e:
        print(f"日期解析错误: {e}")
        return None, None


def click_element_forcefully(driver, element):
    """JS强力点击辅助函数"""
    try:
        driver.execute_script("arguments[0].click();", element)
    except Exception:
        element.click()


def select_dropdown_option(driver, wait, placeholder_text, target_value):
    """
    操作下拉框 (修复版：只点击可见的选项)
    """
    print(f"      正在选择: {target_value} ...")
    try:
        # 1. 找到输入框并点击，让菜单弹出来
        input_xpath = f"//input[@placeholder='{placeholder_text}']"
        input_ele = wait.until(EC.element_to_be_clickable((By.XPATH, input_xpath)))
        click_element_forcefully(driver, input_ele)
        time.sleep(1)  # 等待菜单弹出动画

        # 2. 关键修复：查找所有包含目标文字的选项，但只点“可见”的那一个
        option_xpath = f"//li[contains(., '{target_value}')]"
        options = driver.find_elements(By.XPATH, option_xpath)

        clicked = False
        for opt in options:
            if opt.is_displayed():
                click_element_forcefully(driver, opt)
                clicked = True
                break

        if not clicked:
            print(f"      ⚠️ 警告：找到了选项但它们似乎都被隐藏了，尝试强制点击最后一个...")
            if options:
                click_element_forcefully(driver, options[-1])

        time.sleep(0.5)
    except Exception as e:
        print(f"      ❌ 选择下拉框失败: {e}")


def start_automation():
    print("\n" + "=" * 50)
    print("🤖 平台自动录入系统 (Selenium)")
    print("说明：自动读取【输出结果】中的Excel文件并上传至网页。")
    print("=" * 50)
    print("正在尝试连接已打开的浏览器...")

    try:
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        print("✅ 成功连接到浏览器！")
    except Exception as e:
        print("❌ 连接失败！请检查以下两点：")
        print("1. 是否已通过【专用快捷方式】打开了Chrome浏览器？")
        print("2. 是否已在浏览器中登录并停留在【食材入库维护】页面？")
        input("按回车键返回主菜单...")
        return

    if not os.path.exists(FOLDER_PATH):
        print(f"❌ 错误：文件夹路径不存在 -> {FOLDER_PATH}")
        print("💡 提示：请先执行功能 [2] 生成入库表格。")
        input("按回车键返回主菜单...")
        return

    file_list = [f for f in os.listdir(FOLDER_PATH) if f.endswith('.xls') or f.endswith('.xlsx')]
    file_list.sort()

    if not file_list:
        print("❌ 文件夹里没有找到 Excel 文件！")
        input("按回车键返回主菜单...")
        return

    print("-" * 50)
    print(f"📂 读取路径: {FOLDER_PATH}")
    print(f"📄 待处理文件: {len(file_list)} 个")
    print("👉 请确保浏览器页面停留在【食材入库维护】。")
    print("-" * 50)

    confirm = input("👉 准备好后，按【y】开始，其他键取消: ").strip().lower()
    if confirm != 'y':
        print("🚫 操作已取消。")
        return

    for index, file_name in enumerate(file_list, 1):
        full_file_path = os.path.join(FOLDER_PATH, file_name)
        target_date = file_name.split('.')[0]
        academic_year, semester = get_academic_info(target_date)

        print(f"\n[{index}/{len(file_list)}] 处理文件: {file_name}")
        print(f"   📅 日期: {target_date} -> 学年: {academic_year} | 学期: {semester}")

        try:
            wait = WebDriverWait(driver, 15)

            # === 1. 顶部筛选 ===
            print("   1. 正在切换学期...")
            select_dropdown_option(driver, wait, "请选择学年", academic_year)
            select_dropdown_option(driver, wait, "请选择学期", semester)

            print("      点击查询...")
            query_btn = driver.find_element(By.XPATH, "//button[contains(., '查询')]")
            click_element_forcefully(driver, query_btn)
            time.sleep(2)

            # === 2. 点击“采购食材录入” ===
            print("   2. 打开录入弹窗...")
            try:
                entry_btn = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(., '采购食材录入')]")
                ))
                click_element_forcefully(driver, entry_btn)
            except TimeoutException:
                print("   ⚠️ 按钮没反应，刷新网页重来...")
                driver.refresh()
                time.sleep(5)
                entry_btn = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(., '采购食材录入')]")
                ))
                click_element_forcefully(driver, entry_btn)

            time.sleep(2)

            # === 3. 填写表单 ===
            print("   3. 填写信息...")
            try:
                dazong_radio = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//label[contains(., '大宗食材')]")
                ))
                click_element_forcefully(driver, dazong_radio)
            except:
                pass
            time.sleep(0.5)

            # 填写日期
            js_force_date = f"""
                var inputs = document.querySelectorAll("input");
                inputs.forEach(function(input) {{
                    var p = input.placeholder;
                    if (p && (p.indexOf('采购日期') > -1 || p.indexOf('入库日期') > -1)) {{
                        input.removeAttribute('readonly');
                        input.value = '{target_date}';
                        input.dispatchEvent(new Event('input', {{ bubbles: true }}));
                        input.dispatchEvent(new Event('change', {{ bubbles: true }}));
                        input.dispatchEvent(new Event('blur', {{ bubbles: true }}));
                    }}
                }});
            """
            driver.execute_script(js_force_date)
            time.sleep(1)

            inherit_no_radio = driver.find_element(By.XPATH,
                                                   "//label[contains(@class,'el-radio')][.//span[text()='否']]")
            click_element_forcefully(driver, inherit_no_radio)

            # === 4. 点击“清单导入” ===
            print("   4. 打开导入窗口...")
            time.sleep(1)
            import_btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(., '清单导入')]")
            ))
            click_element_forcefully(driver, import_btn)

            # === 5. 上传文件 ===
            print("   5. 正在上传文件 (等待5秒)...")
            upload_input = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//div[@aria-label='清单导入']//input[@type='file']")
            ))
            upload_input.send_keys(full_file_path)

            time.sleep(5)

            # === 6. 点击“清单导入”弹窗的“确定” ===
            print("   6. 确认导入...")
            try:
                confirm_import_btn = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//div[@aria-label='清单导入']//button[contains(., '确')]")
                ))
                click_element_forcefully(driver, confirm_import_btn)
            except Exception:
                all_confirm_btns = driver.find_elements(By.XPATH, "//button[contains(., '确')]")
                if all_confirm_btns:
                    click_element_forcefully(driver, all_confirm_btns[-1])

            print("      等待数据回填 (3秒)...")
            time.sleep(3)

            # === 7. 点击主界面的“确定”保存 ===
            print("   7. 保存并提交...")
            final_confirm_btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH,
                 "//div[@aria-label='食材入库维护']//div[contains(@class, 'dialog-footer')]//button[contains(., '确')]")
            ))
            driver.execute_script("arguments[0].scrollIntoView();", final_confirm_btn)
            time.sleep(1)
            click_element_forcefully(driver, final_confirm_btn)

            print(f"   ✅ {target_date} 录入成功！")
            print("   🛌 休息4秒...")
            time.sleep(4)

        except Exception as e:
            print(f"❌ ERROR: 处理 {file_name} 时出错!")
            print(f"   错误信息: {e}")
            input("   👉 请手动纠正后按回车继续...")

    print("\n" + "=" * 50)
    print("🎉 所有文件处理完毕！")
    print("=" * 50)
    input("按回车键返回主菜单...")


if __name__ == "__main__":
    start_automation()