import sys
import os

# 将 src 目录添加到路径中，确保可以直接从根目录运行
sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__))))

from src.ui.main_menu import main

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n👋 程序已终止。")
