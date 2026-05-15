import sys
import time
from src.ui.student_cli import run_student_manager
from src.ui.inventory_cli import run_inventory_manager
from src.automation.nutrition_bot import start_automation
from src.utils import ui_utils

def main():
    while True:
        ui_utils.print_banner("🍱 校园营养餐综合管理工具箱", "版本: v2.0 | 模块化重构版")
        print("\n请选择要执行的功能：\n")
        print("  [1] 🎓 学生名单核算 (人数核对、跨班调剂)")
        print("  [2] 🥦 食材入库生成 (自动拆分每日入库单)")
        print("  [3] 🤖 平台自动录入 (Selenium 自动化上传)")
        print("  [0] ❌ 退出系统")
        print("\n" + "-" * 70)
        
        choice = ui_utils.get_input("👉 请输入功能编号", "0")

        if choice == '1':
            run_student_manager()
        elif choice == '2':
            run_inventory_manager()
        elif choice == '3':
            start_automation()
        elif choice == '0':
            print("\n👋 感谢使用，再见！")
            sys.exit()
        else:
            print("\n⚠️ 输入无效，请重新输入...")
            time.sleep(1)
