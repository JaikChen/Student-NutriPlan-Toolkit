import sys
import shutil
import pandas as pd
from datetime import datetime
from src.utils import config, ui_utils
from src.core import student_logic

def handle_old_file(file_path):
    if not file_path.exists():
        return True

    ui_utils.print_banner("⚠️ 文件冲突", f"检测到已存在旧文件: {file_path.name}")
    print("请选择处理方式：")
    print("  [1] 🗑️  删除旧文件 (覆盖)")
    print("  [2] 📦 归档并备份")
    print("  [3] ❌ 取消操作")

    choice = ui_utils.get_input("👉 请输入选择", "2")
    if choice == '1':
        try:
            file_path.unlink()
            return True
        except Exception as e:
            print(f"❌ 删除失败: {e}")
            return False
    elif choice == '2':
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            new_name = f"{file_path.stem}_备份_{timestamp}{file_path.suffix}"
            dest_path = config.STUDENT_ARCHIVE_DIR / new_name
            shutil.move(str(file_path), str(dest_path))
            return True
        except Exception as e:
            print(f"❌ 归档失败: {e}")
            return False
    return False

def format_class_name(raw_grade, raw_class, grade_map):
    g_name = grade_map.get(raw_grade, {}).get('name', str(raw_grade))
    c_num = student_logic.extract_number(raw_class)
    c_name = f"{c_num}班" if c_num > 0 else str(raw_class)
    return f"{g_name} {c_name}"

def run_student_manager():
    ui_utils.print_banner("🎓 学生名单智能核算系统", "智能排序 | 跨班调剂 | 变动日志")
    config.ensure_dirs()

    if not config.STUDENT_INPUT_FILE.exists():
        print(f"\n❌ 未找到源文件: {config.STUDENT_INPUT_FILE}")
        input("按回车键返回...")
        return

    try:
        print("📂 正在读取源文件...")
        df = pd.read_excel(config.STUDENT_INPUT_FILE)
    except Exception as e:
        print(f"❌ 读取失败: {e}")
        input("按回车键返回...")
        return

    grade_map = student_logic.generate_grade_map(df)
    unique_classes = df[['年级', '班级']].drop_duplicates().values.tolist()
    sorted_classes = sorted(unique_classes, key=lambda x: student_logic.get_class_sort_key(x[0], x[1], grade_map))
    total_classes = len(sorted_classes)

    targets_map = {tuple(cls): len(df[(df['年级'] == cls[0]) & (df['班级'] == cls[1])]) for cls in sorted_classes}
    original_counts = targets_map.copy()

    print(f"✅ 读取成功！共 {total_classes} 个班级。")
    
    print("\n请选择录入方式：")
    print("  [1] 📋 批量粘贴")
    print("  [2] ✍️ 逐个输入")
    mode = ui_utils.get_input("👉 模式编号", "1")

    if mode == '1':
        print("\n📢 【批量模式】 请依次输入各班级人数，空格分隔")
        while True:
            clean = input(">> ").replace(',', ' ').replace('，', ' ').replace('\n', ' ')
            try:
                nums = [int(x) for x in clean.split() if x.strip()]
                if len(nums) == total_classes:
                    for idx, cls in enumerate(sorted_classes): targets_map[tuple(cls)] = nums[idx]
                    break
                print(f"⚠️ 数量不匹配 (需{total_classes}, 输{len(nums)})")
            except:
                print("❌ 格式错误。")
    else:
        for cls in sorted_classes:
            name = format_class_name(cls[0], cls[1], grade_map)
            curr = targets_map[tuple(cls)]
            val = input(f"{name:<12} (现{curr}) >> ")
            if val.strip():
                try: targets_map[tuple(cls)] = int(val)
                except: pass

    while True:
        ui_utils.print_banner("🔍 核对清单")
        for idx, cls in enumerate(sorted_classes):
            org = original_counts[tuple(cls)]
            tar = targets_map[tuple(cls)]
            diff = tar - org
            mark = f"{diff:+}" if diff != 0 else "-"
            status = "🔴" if diff < 0 else ("🟢" if diff > 0 else "⚪")
            print(f"{idx + 1:<3} {format_class_name(cls[0], cls[1], grade_map):<10} {org:<4}->{tar:<4} {mark} {status}")

        cmd = input("\n👉 [y]开始 [n]退出 [序号 新值]修改: ").strip().lower()
        if cmd in ['y', '']: break
        if cmd == 'n': return
        try:
            parts = cmd.split()
            t_idx = int(parts[0]) - 1
            if 0 <= t_idx < total_classes:
                new_v = int(parts[1]) if len(parts) > 1 else int(input("新值: "))
                targets_map[tuple(sorted_classes[t_idx])] = new_v
        except: pass

    if handle_old_file(config.STUDENT_OUTPUT_FILE):
        print("\n⏳ 正在计算...")
        final_dfs = []
        all_changes = []
        seen_grades = []
        for g, c in sorted_classes:
            if g not in seen_grades: seen_grades.append(g)

        for grade in seen_grades:
            grade_df = df[df['年级'] == grade]
            processed, logs, changes = student_logic.process_grade_data(grade_df, targets_map, grade)
            final_dfs.extend(processed)
            all_changes.extend(changes)

        if final_dfs:
            result_df = pd.concat(final_dfs)
            change_df = pd.DataFrame(all_changes)
            try:
                with pd.ExcelWriter(config.STUDENT_OUTPUT_FILE) as writer:
                    result_df.to_excel(writer, sheet_name='最终名单', index=False)
                    if not change_df.empty:
                        change_df.to_excel(writer, sheet_name='变动记录', index=False)
                print(f"\n🎉 处理完成！文件已保存至:\n   {config.STUDENT_OUTPUT_FILE}")
            except Exception as e:
                print(f"❌ 保存失败: {e}")

    input("\n按回车键返回主菜单...")
