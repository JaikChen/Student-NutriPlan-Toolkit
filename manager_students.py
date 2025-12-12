import pandas as pd
import re
import os
import sys
import time
import shutil
from datetime import datetime

# ================= 配置区域 =================
BASE_DIR = os.path.join('data', '1_学生名单管理')
INPUT_FILE = os.path.join(BASE_DIR, '营养餐基本名单.xlsx')
OUTPUT_FILE = os.path.join(BASE_DIR, '营养餐_最终核定表.xlsx')
ARCHIVE_DIR = os.path.join(BASE_DIR, '历史备份')  # 新增备份目录


# ===========================================

def init_workspace():
    """初始化工作区"""
    for path in [BASE_DIR, ARCHIVE_DIR]:
        if not os.path.exists(path):
            try:
                os.makedirs(path)
                print(f"✨ 已自动创建文件夹: {path}")
            except Exception as e:
                print(f"❌ 创建文件夹失败: {e}")


def handle_old_file(file_path):
    """处理旧文件冲突"""
    if not os.path.exists(file_path):
        return True  # 没有旧文件，直接通行

    print("\n" + "!" * 50)
    print(f"⚠️  检测到已存在旧文件: {os.path.basename(file_path)}")
    print("请选择处理方式：")
    print("  [1] 🗑️  删除旧文件 (覆盖)")
    print("  [2] 📦 归档并备份 (移至 '历史备份' 文件夹)")
    print("  [3] ❌ 取消操作")
    print("!" * 50)

    while True:
        choice = input("👉 请输入选择 (1/2/3): ").strip()

        if choice == '1':
            try:
                os.remove(file_path)
                print("🗑️  旧文件已删除。")
                return True
            except Exception as e:
                print(f"❌ 删除失败: {e} (请检查文件是否被打开)")
                return False

        elif choice == '2':
            try:
                if not os.path.exists(ARCHIVE_DIR):
                    os.makedirs(ARCHIVE_DIR)

                # 生成带时间戳的新文件名
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = os.path.basename(file_path)
                name, ext = os.path.splitext(filename)
                new_name = f"{name}_备份_{timestamp}{ext}"
                dest_path = os.path.join(ARCHIVE_DIR, new_name)

                shutil.move(file_path, dest_path)
                print(f"📦 已归档至: {dest_path}")
                return True
            except Exception as e:
                print(f"❌ 归档失败: {e} (请检查文件是否被打开)")
                return False

        elif choice == '3':
            print("🚫 操作已取消。")
            return False

        else:
            print("输入无效，请重试。")


def print_header():
    os.system('cls' if os.name == 'nt' else 'clear')
    print("=" * 70)
    print(" " * 15 + "🎓 学生名单智能核算系统")
    print(" " * 18 + "智能排序 | 跨班调剂 | 变动日志")
    print("=" * 70)


def extract_number(text):
    match = re.search(r'(\d+)', str(text))
    return int(match.group(1)) if match else 0


def generate_grade_map(df):
    if '年级' not in df.columns: return {}
    unique_grades = df['年级'].dropna().unique()
    grade_data = []
    for g in unique_grades:
        num = extract_number(g)
        if num > 0: grade_data.append({'raw': g, 'num': num})
    if not grade_data: return {}
    max_num = max(item['num'] for item in grade_data)
    mapping = {}
    is_year_format = max_num > 1900
    base_year = max_num if is_year_format else 0
    cn_nums = {1: '一', 2: '二', 3: '三', 4: '四', 5: '五', 6: '六', 7: '七', 8: '八', 9: '九'}
    for item in grade_data:
        level = base_year - item['num'] + 1 if is_year_format else item['num']
        display_name = f"{cn_nums.get(level, str(level))}年级"
        mapping[item['raw']] = {'sort': level, 'name': display_name}
    return mapping


def get_class_sort_key(raw_grade, raw_class, grade_map):
    g_sort = grade_map[raw_grade]['sort'] if raw_grade in grade_map else 999
    c_num = extract_number(raw_class)
    c_sort = c_num if c_num > 0 else 999
    return (g_sort, c_sort)


def format_class_name(raw_grade, raw_class, grade_map):
    g_name = grade_map.get(raw_grade, {}).get('name', str(raw_grade))
    c_num = extract_number(raw_class)
    c_name = f"{c_num}班" if c_num > 0 else str(raw_class)
    return f"{g_name} {c_name}"


def process_grade_data(grade_df, targets_map, grade_key):
    processed_dfs = []
    summary_logs = []
    change_records = []
    classes = grade_df['班级'].unique()
    spare_pool = []
    class_core_data = {}

    # Step 1: 裁员
    for cls in classes:
        full_key = (grade_key, cls)
        cls_df = grade_df[grade_df['班级'] == cls]
        current_count = len(cls_df)
        target = targets_map.get(full_key, current_count)
        if current_count > target:
            keep_df = cls_df.iloc[:target]
            spares_df = cls_df.iloc[target:]
            class_core_data[cls] = keep_df
            for idx, row in spares_df.iterrows():
                row_dict = row.to_dict()
                row_dict['_origin_class'] = cls
                spare_pool.append(row_dict)
            log = {'班级': cls, '原': current_count, '实': target, '状态': f'📉 移出 {current_count - target} 人'}
        else:
            class_core_data[cls] = cls_df
            log = {'班级': cls, '原': current_count, '实': target, '状态': '⚪ 待定'}
        summary_logs.append(log)

    # Step 2: 补员
    for log in summary_logs:
        cls = log['班级']
        target = log['实']
        current_data = class_core_data[cls]
        current_len = len(current_data)
        needed = target - current_len
        final_cls_df = current_data.copy()
        if needed > 0:
            borrowed_rows = []
            actual_borrowed = 0
            while needed > 0 and spare_pool:
                row_dict = spare_pool.pop(0)
                change_records.append({
                    '年级': grade_key, '姓名': row_dict.get('姓名', '未知'),
                    '原班级': row_dict['_origin_class'], '操作': '借调变动',
                    '现班级': cls, '身份证号': row_dict.get('身份证号', '')
                })
                row_dict['班级'] = cls
                del row_dict['_origin_class']
                borrowed_rows.append(row_dict)
                needed -= 1
                actual_borrowed += 1
            if borrowed_rows:
                borrowed_df = pd.DataFrame(borrowed_rows)
                final_cls_df = pd.concat([final_cls_df, borrowed_df], ignore_index=True)
            if needed == 0:
                log['状态'] = f'📈 借入 {actual_borrowed} 人'
            else:
                log['状态'] = f'⚠️ 借入 {actual_borrowed} (仍缺{needed})'
        elif log['状态'] == '⚪ 待定':
            log['状态'] = '✅ 无变化'
        processed_dfs.append(final_cls_df)

    # Step 3: 删除
    for row_dict in spare_pool:
        change_records.append({
            '年级': grade_key, '姓名': row_dict.get('姓名', '未知'),
            '原班级': row_dict['_origin_class'], '操作': '彻底删除',
            '现班级': '无', '身份证号': row_dict.get('身份证号', '')
        })
    return processed_dfs, summary_logs, change_records


def run_student_manager():
    print_header()
    init_workspace()

    if not os.path.exists(INPUT_FILE):
        print(f"\n❌ 未找到源文件: {INPUT_FILE}")
        print("💡 请将 Excel 文件放入文件夹后重试。")
        input("按回车键返回...")
        return

    try:
        print("📂 正在读取源文件...")
        df = pd.read_excel(INPUT_FILE)
    except Exception as e:
        print(f"❌ 读取失败: {e}")
        input("按回车键返回...")
        return

    grade_map = generate_grade_map(df)
    unique_classes = df[['年级', '班级']].drop_duplicates().values.tolist()
    sorted_classes = sorted(unique_classes, key=lambda x: get_class_sort_key(x[0], x[1], grade_map))
    total_classes = len(sorted_classes)

    targets_map = {}
    original_counts = {}
    for g, c in sorted_classes:
        curr = len(df[(df['年级'] == g) & (df['班级'] == c)])
        original_counts[(g, c)] = curr
        targets_map[(g, c)] = curr

    print(f"✅ 读取成功！共 {total_classes} 个班级。")
    time.sleep(0.5)

    print("\n请选择录入方式：")
    print("  [1] 📋 批量粘贴")
    print("  [2] ✍️ 逐个输入")
    while True:
        mode = input("\n👉 模式编号: ").strip()
        if mode in ['1', '2']: break

    if mode == '1':
        print("\n📢 【批量模式】")
        print(f"顺序: {format_class_name(sorted_classes[0][0], sorted_classes[0][1], grade_map)} ...")
        while True:
            clean = input(">> ").replace(',', ' ').replace('，', ' ').replace('\n', ' ')
            try:
                nums = [int(x) for x in clean.split() if x.strip()]
                if len(nums) == total_classes:
                    for idx, (g, c) in enumerate(sorted_classes): targets_map[(g, c)] = nums[idx]
                    break
                else:
                    print(f"⚠️ 数量不匹配 (需{total_classes}, 输{len(nums)})")
            except:
                print("❌ 格式错误。")
    else:
        print("\n📢 【逐个模式】回车跳过")
        for g, c in sorted_classes:
            name = format_class_name(g, c, grade_map)
            curr = targets_map[(g, c)]
            val = input(f"{name:<12} (现{curr}) >> ")
            if val.strip():
                try:
                    targets_map[(g, c)] = int(val)
                except:
                    pass

    while True:
        os.system('cls' if os.name == 'nt' else 'clear')
        print(f"\n🔍 核对清单")
        diff_total = 0
        for idx, (g, c) in enumerate(sorted_classes):
            org = original_counts[(g, c)]
            tar = targets_map[(g, c)]
            diff = tar - org
            mark = f"{diff:+}" if diff != 0 else "-"
            status = "🔴" if diff < 0 else ("🟢" if diff > 0 else "⚪")
            print(f"{idx + 1:<3} {format_class_name(g, c, grade_map):<10} {org:<4}->{tar:<4} {mark} {status}")
            diff_total += tar

        print("-" * 50)
        cmd = input("👉 [y]开始 [n]退出 [序号 新值]修改: ").strip().lower()
        if cmd == 'y' or cmd == '': break
        if cmd == 'n': return
        parts = cmd.split()
        if len(parts) >= 1:
            try:
                t_idx = int(parts[0]) - 1
                if 0 <= t_idx < total_classes:
                    new_v = int(parts[1]) if len(parts) > 1 else int(input("新值: "))
                    targets_map[sorted_classes[t_idx]] = new_v
            except:
                pass

    # ================= 核心修改：保存前的冲突检测 =================

    # 在计算前先确认用户是否想继续（如果旧文件处理失败，这里就不必计算了）
    if os.path.exists(OUTPUT_FILE):
        if not handle_old_file(OUTPUT_FILE):
            input("按回车键返回...")
            return

    print("\n⏳ 正在计算...")
    final_dfs = []
    all_changes = []

    sorted_grades = []
    seen = set()
    for g, c in sorted_classes:
        if g not in seen: sorted_grades.append(g); seen.add(g)

    for grade in sorted_grades:
        grade_df = df[df['年级'] == grade]
        processed, logs, changes = process_grade_data(grade_df, targets_map, grade)
        final_dfs.extend(processed)
        all_changes.extend(changes)

    if final_dfs:
        result_df = pd.concat(final_dfs)
        change_df = pd.DataFrame(all_changes)

        try:
            with pd.ExcelWriter(OUTPUT_FILE) as writer:
                result_df.to_excel(writer, sheet_name='最终名单', index=False)
                if not change_df.empty:
                    change_df.to_excel(writer, sheet_name='变动记录', index=False)
                else:
                    pd.DataFrame({'提示': ['无变动']}).to_excel(writer, sheet_name='变动记录', index=False)
            print(f"\n🎉 处理完成！文件已保存至:\n   {OUTPUT_FILE}")
        except Exception as e:
            print(f"❌ 保存失败: {e}")

    input("\n按回车键返回主菜单...")


if __name__ == "__main__":
    run_student_manager()