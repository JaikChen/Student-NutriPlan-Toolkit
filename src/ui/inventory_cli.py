import shutil
import pandas as pd
from datetime import datetime
from src.utils import config, ui_utils
from src.core import inventory_logic

def handle_existing_outputs():
    files = list(config.INVENTORY_OUTPUT_DIR.glob('*.xls'))
    if not files:
        return True

    ui_utils.print_banner("⚠️ 输出目录冲突", f"检测到输出目录中已有 {len(files)} 个文件。")
    print("请选择处理方式：")
    print("  [1] 🗑️  清空输出目录")
    print("  [2] 📦 归档当前文件")
    print("  [3] 🐢 保留旧文件")
    print("  [4] ❌ 取消操作")
    
    choice = ui_utils.get_input("👉 请输入选择", "2")
    if choice == '1':
        for f in files: f.unlink()
        return True
    elif choice == '2':
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest_path = config.INVENTORY_ARCHIVE_DIR / f"备份_{timestamp}"
        dest_path.mkdir(parents=True, exist_ok=True)
        for f in files: shutil.move(str(f), str(dest_path / f.name))
        return True
    elif choice == '3':
        return True
    return False

def run_inventory_manager():
    ui_utils.print_banner("🥦 食材入库单生成工具", "读取采购清单并按日期拆分模板")
    config.ensure_dirs()

    if not config.INVENTORY_INPUT_FILE.exists() or not config.INVENTORY_TEMPLATE_FILE.exists():
        print(f"\n❌ 缺少必要文件，请确保以下文件存在:\n   - {config.INVENTORY_INPUT_FILE}\n   - {config.INVENTORY_TEMPLATE_FILE}")
        input("按回车键返回...")
        return

    try:
        df = pd.read_excel(config.INVENTORY_INPUT_FILE, header=1)
        df.columns = df.columns.str.strip()
    except Exception as e:
        print(f"❌ 读取失败: {e}")
        input("按回车键返回...")
        return

    if '采购日期' not in df.columns:
        print("❌ 错误：未找到 '采购日期' 列。")
        input("按回车键返回...")
        return

    if not handle_existing_outputs():
        return

    print("\n⚡ 开始处理...")
    count, errors = inventory_logic.generate_inventory_files(
        df, config.INVENTORY_TEMPLATE_FILE, config.INVENTORY_OUTPUT_DIR
    )

    for date, err in errors:
        print(f"   ❌ 日期 {date} 处理失败: {err}")

    print("\n" + "=" * 50)
    print(f"🎉 全部完成！共生成 {count} 个文件。")
    print(f"📂 输出位置: {config.INVENTORY_OUTPUT_DIR}")
    input("按回车键返回主菜单...")
