import pandas as pd
import os
import xlrd
from xlutils.copy import copy
import datetime

# =================配置区域=================
PARENT_FOLDER = '食材入库管理'
INPUT_FOLDER = os.path.join(PARENT_FOLDER, '1_把源文件放这里')
OUTPUT_FOLDER = os.path.join(PARENT_FOLDER, '2_生成的上传文件')

SOURCE_FILE_NAME = '采购清单.xlsx'  # 数据源
TEMPLATE_FILE_NAME = '食材入库信息表.xls'  # 必须是原版 .xls 模板


# =========================================

def process_xls_template():
    print(f"🚀 启动【.xls 原版模板填充模式】...")

    # 1. 检查文件夹和文件
    if not os.path.exists(INPUT_FOLDER):
        print(f"❌ 文件夹不存在: {INPUT_FOLDER}")
        return

    source_path = os.path.join(INPUT_FOLDER, SOURCE_FILE_NAME)
    template_path = os.path.join(INPUT_FOLDER, TEMPLATE_FILE_NAME)

    if not os.path.exists(source_path):
        print(f"❌ 缺少数据源: {SOURCE_FILE_NAME}")
        return
    if not os.path.exists(template_path):
        print(f"❌ 缺少模板文件: {TEMPLATE_FILE_NAME}")
        print("👉 请把平台下载的原始 .xls 文件放进去！")
        return

    # 2. 读取数据源
    print(f"📖 读取数据源...")
    try:
        # header=1 跳过第一行日期，从第二行开始读表头
        df = pd.read_excel(source_path, header=1)
        df.columns = df.columns.str.strip()
    except Exception as e:
        print(f"❌ 数据源读取失败: {e}")
        return

    # 3. 准备输出目录
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)

    # 4. 按日期拆分并填充
    grouped = df.groupby('采购日期')
    target_columns = ["食材名称", "食材单位", "食材数量", "食材单价", "小计"]
    count = 0

    print("⚡ 开始生成 .xls 文件...")

    for date, group in grouped:
        try:
            # A. 打开原版模板 (启用 formatting_info=True 以保留格式)
            rb = xlrd.open_workbook(template_path, formatting_info=True)

            # B. 复制一个可写入的副本
            wb = copy(rb)
            ws = wb.get_sheet(0)  # 获取第一个工作表

            # C. 准备写入的数据
            upload_data = group[target_columns].copy()

            # D. 写入数据 (从第2行索引开始，即视觉上的第3行)
            # 模板结构：Row 0 = 标题, Row 1 = 表头, Row 2 = 数据开始
            start_row = 2

            # 遍历数据写入
            for r_idx, (index, row) in enumerate(upload_data.iterrows()):
                # row 是一个 Series，包含那一行的5列数据
                # 写入 5 列: 名称(0), 单位(1), 数量(2), 单价(3), 小计(4)
                ws.write(start_row + r_idx, 0, row['食材名称'])
                ws.write(start_row + r_idx, 1, row['食材单位'])
                ws.write(start_row + r_idx, 2, row['食材数量'])
                ws.write(start_row + r_idx, 3, row['食材单价'])
                ws.write(start_row + r_idx, 4, row['小计'])

            # E. 保存为 .xls 文件
            date_str = str(date).split(' ')[0]
            save_filename = f"{date_str}.xls"  # 保持 .xls 后缀
            save_path = os.path.join(OUTPUT_FOLDER, save_filename)

            wb.save(save_path)
            print(f"   ✅ 已生成: {save_filename}")
            count += 1

        except Exception as e:
            print(f"   ❌ 处理日期 {date} 时出错: {e}")

    print("\n" + "=" * 40)
    print(f"🎉 全部完成！共生成 {count} 个标准 .xls 文件。")
    print(f"📂 请直接上传此文件夹内的文件: {OUTPUT_FOLDER}")
    print("=" * 40)


if __name__ == "__main__":
    process_xls_template()