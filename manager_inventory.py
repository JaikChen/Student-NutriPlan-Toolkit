import pandas as pd
import os
import shutil
import xlrd
from xlutils.copy import copy
import datetime

# ================= 配置区域 =================
BASE_DIR = os.path.join('data', '2_食材入库管理')
INPUT_FILE = os.path.join(BASE_DIR, '采购清单.xlsx')
TEMPLATE_FILE = os.path.join(BASE_DIR, '食材入库信息表.xls')
OUTPUT_DIR = os.path.join(BASE_DIR, '输出结果')
ARCHIVE_DIR = os.path.join(BASE_DIR, '历史备份')


# ===========================================

def init_workspace():
    """初始化工作区"""
    for path in [BASE_DIR, OUTPUT_DIR, ARCHIVE_DIR]:
        if not os.path.exists(path):
            os.makedirs(path)
            print(f"✨ 已自动创建文件夹: {path}")


def handle_existing_outputs():
    """处理已存在的输出文件"""
    # 检查输出目录是否有 .xls 文件
    files = [f for f in os.listdir(OUTPUT_DIR) if f.endswith('.xls')]
    if not files:
        return True  # 目录是空的，直接继续

    print("\n" + "!" * 50)
    print(f"⚠️  检测到输出目录 '{os.path.basename(OUTPUT_DIR)}' 中已有 {len(files)} 个文件。")
    print("为避免混淆，建议先清理旧文件。请选择：")
    print("  [1] 🗑️  清空输出目录 (删除所有旧 .xls)")
    print("  [2] 📦 归档当前文件 (移至 '历史备份')")
    print("  [3] 🐢 保留旧文件 (新文件将直接混入/覆盖)")
    print("  [4] ❌ 取消操作")
    print("!" * 50)

    while True:
        choice = input("👉 请输入选择 (1/2/3/4): ").strip()

        if choice == '1':
            try:
                for f in files:
                    os.remove(os.path.join(OUTPUT_DIR, f))
                print("🗑️  目录已清空。")
                return True
            except Exception as e:
                print(f"❌ 清空失败: {e}")
                return False

        elif choice == '2':
            try:
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_folder_name = f"入库单备份_{timestamp}"
                dest_path = os.path.join(ARCHIVE_DIR, backup_folder_name)

                os.makedirs(dest_path)

                for f in files:
                    shutil.move(os.path.join(OUTPUT_DIR, f), os.path.join(dest_path, f))

                print(f"📦 已将 {len(files)} 个文件移动至: {dest_path}")
                return True
            except Exception as e:
                print(f"❌ 归档失败: {e}")
                return False

        elif choice == '3':
            print("🐢 保持现状，继续生成...")
            return True

        elif choice == '4':
            print("🚫 操作已取消。")
            return False
        else:
            print("输入无效。")


def run_inventory_manager():
    print("\n" + "=" * 50)
    print("🥦 食材入库单生成工具")
    print("说明：读取 '采购清单.xlsx'，按日期拆分并填充到 '.xls' 模板中。")
    print("=" * 50)

    init_workspace()

    # 1. 检查必要文件
    if not os.path.exists(INPUT_FILE) or not os.path.exists(TEMPLATE_FILE):
        print(f"\n❌ 缺少文件，请检查: {BASE_DIR}")
        input("按回车键返回...")
        return

    # 2. 读取数据
    print(f"📖 正在读取采购清单...")
    try:
        df = pd.read_excel(INPUT_FILE, header=1)
        df.columns = df.columns.str.strip()
    except Exception as e:
        print(f"❌ 读取失败: {e}")
        input("按回车键返回...")
        return

    if '采购日期' not in df.columns:
        print("❌ 错误：表格中未找到 '采购日期' 列。")
        input("按回车键返回...")
        return

    # 3. 处理旧文件 (核心更新)
    if not handle_existing_outputs():
        return

    grouped = df.groupby('采购日期')
    target_columns = ["食材名称", "食材单位", "食材数量", "食材单价", "小计"]

    count = 0
    print("\n⚡ 开始处理...")

    for date, group in grouped:
        try:
            rb = xlrd.open_workbook(TEMPLATE_FILE, formatting_info=True)
            wb = copy(rb)
            ws = wb.get_sheet(0)

            upload_data = group[target_columns].copy()
            start_row = 2

            for r_idx, (_, row) in enumerate(upload_data.iterrows()):
                ws.write(start_row + r_idx, 0, row['食材名称'])
                ws.write(start_row + r_idx, 1, row['食材单位'])
                ws.write(start_row + r_idx, 2, row['食材数量'])
                ws.write(start_row + r_idx, 3, row['食材单价'])
                ws.write(start_row + r_idx, 4, row['小计'])

            date_str = str(date).split(' ')[0]
            save_path = os.path.join(OUTPUT_DIR, f"{date_str}.xls")

            wb.save(save_path)
            print(f"   ✅ 生成: {date_str}.xls")
            count += 1

        except Exception as e:
            print(f"   ❌ 日期 {date} 处理失败: {e}")

    print("\n" + "=" * 50)
    print(f"🎉 全部完成！共生成 {count} 个文件。")
    print(f"📂 输出位置: {OUTPUT_DIR}")
    input("按回车键返回主菜单...")


if __name__ == "__main__":
    run_inventory_manager()