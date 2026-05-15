import time
import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from src.utils import config, ui_utils

def get_academic_info(date_str):
    """根据日期判断学年和学期"""
    try:
        date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        year, month = date_obj.year, date_obj.month
        if 2 <= month <= 8: return f"{year - 1}-{year}", "春季学期"
        return (f"{year}-{year + 1}", "秋季学期") if month >= 9 else (f"{year - 1}-{year}", "秋季学期")
    except Exception as e:
        print(f"⚠️ 日期解析错误: {e}")
        return None, None

def click_force(driver, element):
    """JS 强力点击"""
    try:
        driver.execute_script("arguments[0].click();", element)
    except:
        element.click()

def find_target_tab(driver):
    """定位目标标签页"""
    keywords = ["营养", "采购", "食材", "管理系统"]
    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        try:
            if any(k in driver.title for k in keywords): return True
        except: pass
    return False

def select_option(driver, wait, placeholder, value):
    """操作 ElementUI 下拉框"""
    try:
        xpath = f"//input[@placeholder='{placeholder}']"
        input_ele = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        click_force(driver, input_ele)
        time.sleep(0.8)
        opt_xpath = f"//li[contains(., '{value}') and contains(@class, 'el-select-dropdown__item')]"
        option = wait.until(EC.visibility_of_element_located((By.XPATH, opt_xpath)))
        click_force(driver, option)
        time.sleep(0.5)
        return True
    except Exception as e:
        print(f"   ⚠️ 选择 {value} 失败: {e}")
        return False

def start_automation():
    ui_utils.print_banner("🤖 平台自动录入机器人", "基于 Selenium 的自动化上传工具")
    config.ensure_dirs()

    print("🚀 正在初始化浏览器引擎...")
    options = Options()
    options.add_argument(f"user-data-dir={config.CHROME_PROFILE_DIR}")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        driver.maximize_window()
        driver.get(config.TARGET_URL)
    except Exception as e:
        print(f"❌ 浏览器启动失败: {e}")
        input("按回车键返回...")
        return

    print("\n" + "!" * 50)
    print("👉 请在浏览器中完成登录")
    print("👉 登录后点击进入【采购管理】 -> 【食材入库维护】")
    print("!" * 50 + "\n")

    # 智能轮询等待进入目标页面
    while True:
        try:
            if find_target_tab(driver):
                if "食材入库维护" in driver.page_source or "procurementStorage" in driver.current_url:
                    print("🎯 已进入目标页面，开始执行任务...")
                    break
        except: pass
        time.sleep(2)

    files = sorted(list(config.INVENTORY_OUTPUT_DIR.glob('*.xls*')))
    if not files:
        print("❌ 待上传文件夹为空，请先执行功能 [2] 生成入库单。")
        input("按回车键返回...")
        return

    print(f"📂 发现 {len(files)} 个待处理文件\n")

    for i, file_path in enumerate(files, 1):
        date_str = file_path.stem
        year, semester = get_academic_info(date_str)
        print(f"[{i}/{len(files)}] 正在处理: {date_str} ({year} {semester})")

        try:
            wait = WebDriverWait(driver, 15)
            
            # 1. 筛选
            select_option(driver, wait, "请选择学年", year)
            select_option(driver, wait, "请选择学期", semester)
            
            search_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".yycSearchBtn")))
            click_force(driver, search_btn)
            time.sleep(1.5)

            # 2. 点击录入
            entry_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., '采购食材录入')]")))
            click_force(driver, entry_btn)
            time.sleep(1)

            # 3. 强力填充日期 (JS 注入)
            js_fill = f"""
                document.querySelectorAll('input').forEach(input => {{
                    let p = input.placeholder || '';
                    if(p.includes('采购日期') || p.includes('入库日期')) {{
                        input.removeAttribute('readonly');
                        input.value = '{date_str}';
                        input.dispatchEvent(new Event('input', {{ bubbles: true }}));
                        input.dispatchEvent(new Event('change', {{ bubbles: true }}));
                    }}
                }});
            """
            driver.execute_script(js_fill)
            time.sleep(0.5)

            # 4. 导入清单
            import_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., '清单导入')]")))
            click_force(driver, import_btn)
            
            file_input = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@aria-label='清单导入']//input[@type='file']")))
            file_input.send_keys(str(file_path.absolute()))
            time.sleep(3) # 等待上传处理

            # 5. 确认保存
            confirm_import = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@aria-label='清单导入']//button[contains(., '确')]")))
            click_force(driver, confirm_import)
            time.sleep(1.5)

            save_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@aria-label='食材入库维护']//button[contains(., '确')]")))
            click_force(driver, save_btn)
            
            print(f"   ✅ {date_str} 上传成功")
            time.sleep(2)

        except Exception as e:
            print(f"   ❌ 处理出错: {e}")
            if i < len(files):
                ans = ui_utils.get_input("   👉 是否手动修正后继续？(y/n)", "y")
                if ans.lower() != 'y': break

    print("\n🎉 自动化任务执行完毕！")
    input("按回车键返回...")
