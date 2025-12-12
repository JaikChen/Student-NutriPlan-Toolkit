import os
import sys
import time
# 确保能导入同目录下的模块
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from manager_students import run_student_manager
from manager_inventory import run_inventory_manager

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def print_main_menu():
    clear_screen()
    print("=" * 60)
    print(" " * 12 + "🍱 校园营养餐综合管理工具箱")
    print("=" * 60)
    print("\n请选择要执行的功能：\n")
    print("  [1] 🎓 学生名单核算 (人数核对、跨班调剂)")
    print("  [2] 🥦 食材入库生成 (自动拆分每日入库单)")
    print("  [0] ❌ 退出系统")
    print("-" * 60)

def main():
    while True:
        print_main_menu()
        choice = input("👉 请输入功能编号: ").strip()

        if choice == '1':
            run_student_manager()
        elif choice == '2':
            run_inventory_manager()
        elif choice == '0':
            print("\n👋 感谢使用，再见！")
            sys.exit()
        else:
            print("\n⚠️ 输入无效，请重新输入...")
            time.sleep(1)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n👋 程序已终止。")