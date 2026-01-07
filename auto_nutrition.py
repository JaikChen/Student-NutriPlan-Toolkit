import time
import os
import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException

# ================= 配置区域 =================
# 获取当前脚本所在目录
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
# 自动定位到 manager_inventory.py 生成结果的目录
FOLDER_PATH = os.path.join(CURRENT_DIR, 'data', '2_食材入库管理', '输出结果')


# ===========================================

def get_academic_info(date_str):
    """根据日期判断学年和学期"""
    try:
        date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        year = date_obj.year
        month = date_obj.month
        # 逻辑：2-8月是春季，9-次年1月是秋季
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
    """JS强力点击"""
    try:
        driver.execute_script("arguments[0].click();", element)
    except Exception:
        element.click()


def switch_to_target_tab(driver):
    """锁定目标标签页"""
    print("🔄 正在扫描并锁定正确的浏览器标签页...")
    target_keywords = ["营养", "采购", "食材", "管理系统"]

    # 1. 检查当前页
    try:
        if any(k in driver.title for k in target_keywords):
            print(f"✅ 已锁定当前页面: {driver.title}")
            return True
    except:
        pass

    # 2. 遍历切换
    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        time.sleep(0.2)
        if any(k in driver.title for k in target_keywords):
            print(f"   🎯 成功切换到目标页面: {driver.title}")
            return True

    print("❌ 警告：未找到包含'营养/采购'字样的标签页！")
    return False


def select_dropdown_option(driver, wait, placeholder_text, target_value):
    """操作下拉框"""
    print(f"      正在选择: {target_value} ...")
    try:
        input_xpath = f"//input[@placeholder='{placeholder_text}']"
        input_ele = wait.until(EC.presence_of_element_located((By.XPATH, input_xpath)))

        try:
            driver.execute_script("arguments[0].parentNode.click();", input_ele)
        except:
            click_element_forcefully(driver, input_ele)
        time.sleep(1)

        option_xpath = f"//li[contains(., '{target_value}')]"
        # 循环尝试点击可见选项
        for _ in range(3):
            options = driver.find_elements(By.XPATH, option_xpath)
            for opt in options:
                if opt.is_displayed():
                    click_element_forcefully(driver, opt)
                    time.sleep(0.5)
                    return
            time.sleep(0.5)

        # 兜底盲点
        options = driver.find_elements(By.XPATH, option_xpath)
        if options: click_element_forcefully(driver, options[-1])
        time.sleep(0.5)
    except Exception:
        pass


def reset_page_state(driver):
    """清理弹窗"""
    try:
        webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(0.5)
    except:
        pass


def start_automation():
    print("\n" + "=" * 50)
    print("🤖 平台自动录入系统 (日期修复版)")
    print("=" * 50)

    print("正在连接浏览器...")
    try:
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        print("✅ 连接成功！")
    except Exception:
        print("❌ 连接失败！请确认已双击【专用快捷方式】打开了浏览器。")
        input("按回车返回...")
        return

    if not switch_to_target_tab(driver):
        print("⚠️ 请手动点击【食材入库维护】页面，然后按回车。")
        input(">>>")

    if not os.path.exists(FOLDER_PATH):
        print(f"❌ 路径不存在: {FOLDER_PATH}")
        return
    file_list = [f for f in os.listdir(FOLDER_PATH) if f.endswith('.xls') or f.endswith('.xlsx')]
    file_list.sort()
    if not file_list:
        print("❌ 文件夹为空。")
        return

    print("-" * 50)
    print(f"📂 待处理文件: {len(file_list)} 个")
    print("👉 请确保浏览器已显示【食材入库维护】界面。")
    input("👉 准备好后，按【回车键】开始 >>> ")

    for index, file_name in enumerate(file_list, 1):
        full_file_path = os.path.join(FOLDER_PATH, file_name)
        target_date = file_name.split('.')[0]
        academic_year, semester = get_academic_info(target_date)

        print(f"\n[{index}/{len(file_list)}] 处理: {file_name} ({academic_year} {semester})")

        try:
            wait = WebDriverWait(driver, 10)
            reset_page_state(driver)

            # === 1. 顶部筛选 ===
            print("   1. 筛选学期...")
            select_dropdown_option(driver, wait, "请选择学年", academic_year)
            select_dropdown_option(driver, wait, "请选择学期", semester)

            print("      点击查询...")
            try:
                query_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".yycSearchBtn")))
                click_element_forcefully(driver, query_btn)
            except:
                try:
                    query_btn = driver.find_element(By.XPATH, "//button[contains(., '查询')]")
                    click_element_forcefully(driver, query_btn)
                except:
                    print("      ⚠️ 查询按钮点击失败，请手动点击...")
            time.sleep(2)

            # === 2. 点击录入 ===
            print("   2. 打开录入...")
            try:
                entry_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., '采购食材录入')]")))
                click_element_forcefully(driver, entry_btn)
            except TimeoutException:
                print("      ❌ 找不到录入按钮！可能是查询未刷新。")
                input("      👉 请手动点击【采购食材录入】，然后按回车...")
            time.sleep(1.5)

            # === 3. 填写表单 ===
            print("   3. 填写表单...")

            try:
                dazong = driver.find_element(By.XPATH, "//label[contains(., '大宗食材')]")
                click_element_forcefully(driver, dazong)
            except:
                pass
            time.sleep(0.5)

            # === 关键修复：强力日期填充 ===
            # 这里加回了 bubbles: true，这是让 ElementUI 感知到数据变化的关键
            js_date_fix = f"""
                var inputs = document.querySelectorAll("input");
                var target = '{target_date}';
                for(var i=0; i<inputs.length; i++) {{
                    var p = inputs[i].placeholder;
                    if(p && (p.indexOf('采购日期')>-1 || p.indexOf('入库日期')>-1)) {{
                        // 1. 移除只读
                        inputs[i].removeAttribute('readonly');
                        // 2. 赋值
                        inputs[i].value = target;
                        // 3. 触发全套事件 (必须加 bubbles: true)
                        inputs[i].dispatchEvent(new Event('input', {{ bubbles: true }}));
                        inputs[i].dispatchEvent(new Event('change', {{ bubbles: true }}));
                        inputs[i].dispatchEvent(new Event('blur', {{ bubbles: true }}));
                    }}
                }}
            """
            driver.execute_script(js_date_fix)
            # 再次确认：有时候JS执行太快，稍微等一下再执行一次保险
            time.sleep(0.2)
            driver.execute_script(js_date_fix)

            # 选否
            try:
                no_btn = driver.find_element(By.XPATH, "//label[contains(@class,'el-radio')][.//span[text()='否']]")
                click_element_forcefully(driver, no_btn)
            except:
                pass

            # === 4. 导入文件 ===
            print("   4. 导入文件...")
            time.sleep(1)
            try:
                import_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., '清单导入')]")))
                click_element_forcefully(driver, import_btn)
            except:
                input("      ⚠️ 找不到【清单导入】按钮，请手动点击后回车...")

            # 上传
            print("      上传中...")
            try:
                file_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//div[@aria-label='清单导入']//input[@type='file']")))
                file_input.send_keys(full_file_path)
            except:
                print("      ❌ 无法定位上传框，请手动上传。")

            time.sleep(4)

            # === 5. 确认 ===
            print("   5. 确认保存...")
            try:
                confirm_import = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//div[@aria-label='清单导入']//button[contains(., '确')]")))
                click_element_forcefully(driver, confirm_import)
            except:
                pass

            time.sleep(2)

            try:
                final_confirm = wait.until(EC.element_to_be_clickable((By.XPATH,
                                                                       "//div[@aria-label='食材入库维护']//div[contains(@class, 'dialog-footer')]//button[contains(., '确')]")))
                click_element_forcefully(driver, final_confirm)
            except:
                pass

            print(f"   ✅ {target_date} 完成！")
            time.sleep(3)

        except Exception as e:
            print(f"❌ 发生异常: {e}")
            input("👉 请手动修正，按回车继续...")

    print("\n🎉 全部完成！")
    input("按回车退出...")


if __name__ == "__main__":
    start_automation()