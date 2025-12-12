import pandas as pd
import re
import os
import sys
import time


# ===========================
# 1. 界面美化与工具模块
# ===========================

def print_header():
    """打印漂亮的程序头"""
    os.system('cls' if os.name == 'nt' else 'clear')  # 清屏
    print("=" * 70)
    print(" " * 15 + "🏫 营养餐名单智能管理系统 (终极版)")
    print(" " * 18 + "智能排序 | 跨班调剂 | 变动日志")
    print("=" * 70)
    print("说明：本程序将读取 '营养餐基本名单.xlsx'，并生成核定后的新表格。\n")


def print_section(title):
    """打印章节标题"""
    print(f"\n\n>> {title}")
    print("-" * 50)


def extract_number(text):
    match = re.search(r'(\d+)', str(text))
    return int(match.group(1)) if match else 0


def generate_grade_map(df):
    """智能解析年级逻辑"""
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
        if is_year_format:
            level = base_year - item['num'] + 1
        else:
            level = item['num']
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


# ===========================
# 2. 核心逻辑 (业务处理)
# ===========================

def process_grade_data(grade_df, targets_map, grade_key):
    processed_dfs = []
    summary_logs = []
    change_records = []

    classes = grade_df['班级'].unique()
    spare_pool = []
    class_core_data = {}

    # === Step 1: 裁员 (收集筹码) ===
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

    # === Step 2: 补员 (分配筹码) ===
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

    # === Step 3: 记录删除 ===
    for row_dict in spare_pool:
        change_records.append({
            '年级': grade_key, '姓名': row_dict.get('姓名', '未知'),
            '原班级': row_dict['_origin_class'], '操作': '彻底删除',
            '现班级': '无', '身份证号': row_dict.get('身份证号', '')
        })

    return processed_dfs, summary_logs, change_records


# ===========================
# 3. 主程序入口
# ===========================

def main():
    print_header()

    input_file = '营养餐基本名单.xlsx'
    output_file = '营养餐_最终核定表.xlsx'

    # --- 1. 智能文件检查 ---
    if not os.path.exists(input_file):
        print(f"❌ 错误：在当前目录下找不到 '{input_file}'")
        print("\n当前目录下的文件有：")
        files = [f for f in os.listdir('.') if f.endswith('.xlsx')]
        if files:
            for f in files: print(f" - {f}")
        else:
            print(" (当前目录没有Excel文件)")
        print("\n💡 建议：请把名单重命名为 '营养餐基本名单.xlsx' 后重新运行。")
        input("按回车键退出...");
        return

    try:
        print("📂 正在读取源文件...")
        df = pd.read_excel(input_file)
    except Exception as e:
        print(f"❌ 读取文件失败: {e}")
        print("请检查文件是否被其他程序占用。")
        input("按回车键退出...");
        return

    # --- 2. 数据初始化 ---
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

    print(f"✅ 读取成功！检测到 {total_classes} 个班级。")
    time.sleep(0.5)

    # --- 3. 交互式输入 ---
    print_section("数据录入")
    print("请选择一种录入方式：")
    print("  [1] 📋 批量粘贴 (推荐：复制一整行数字)")
    print("  [2] ✍️ 逐个输入 (按班级顺序逐个核对)")

    while True:
        mode = input("\n👉 请输入模式编号 (1/2): ").strip()
        if mode in ['1', '2']: break
        print("输入错误，请输入 1 或 2。")

    if mode == '1':
        print("\n📢 【批量模式提示】")
        print("系统识别的班级顺序如下：")
        first = format_class_name(sorted_classes[0][0], sorted_classes[0][1], grade_map)
        last = format_class_name(sorted_classes[-1][0], sorted_classes[-1][1], grade_map)
        print(f"   {first}  ---> ... --->  {last}")
        print("-" * 30)
        print("请直接粘贴人数数字串 (用空格、逗号或换行分隔均可)：")

        while True:
            raw = input(">> ")
            clean = raw.replace(',', ' ').replace('，', ' ').replace('\n', ' ')
            try:
                nums = [int(x) for x in clean.split() if x.strip()]
            except:
                print("❌ 内容包含非数字字符，请重新粘贴。")
                continue

            if len(nums) == total_classes:
                for idx, (g, c) in enumerate(sorted_classes): targets_map[(g, c)] = nums[idx]
                print("✅ 格式正确，录入完成。")
                break
            else:
                print(f"⚠️ 数量不匹配！系统检测到 {total_classes} 个班，您输入了 {len(nums)} 个数字。")
                print("请检查是否漏输，并重新粘贴。")

    else:
        print("\n📢 【逐个模式提示】直接回车代表人数不变。")
        for g, c in sorted_classes:
            name = format_class_name(g, c, grade_map)
            curr = targets_map[(g, c)]
            while True:
                val = input(f"{name:<12} (现有 {curr} 人) >> 实际: ")
                if not val.strip(): break
                try:
                    targets_map[(g, c)] = int(val)
                    break
                except:
                    print("请输入有效数字。")

    # --- 4. 仪表盘式核对清单 (核心人性化升级) ---
    while True:
        os.system('cls' if os.name == 'nt' else 'clear')
        print_section("🔍 核对清单 (Dashboard)")
        print(f"{'No.':<4} {'班级名称':<12} {'原人数':<6} {'新人数':<6} {' 差额':<6} {'状态'}")
        print("-" * 65)

        diff_total = 0
        org_total = sum(original_counts.values())

        for idx, (g, c) in enumerate(sorted_classes):
            name = format_class_name(g, c, grade_map)
            org = original_counts[(g, c)]
            tar = targets_map[(g, c)]
            diff = tar - org

            # 视觉标记
            if diff < 0:
                status = "🔴 删减"
                diff_str = str(diff)
            elif diff > 0:
                status = "🟢 需借"
                diff_str = f"+{diff}"
            else:
                status = "⚪"
                diff_str = "-"

            # 高亮显示有变动的行
            line = f"{idx + 1:<4} {name:<12} {org:<6} {tar:<6} {diff_str:<6} {status}"
            print(line)
            diff_total += tar

        print("-" * 65)
        print(f"【合计】 原: {org_total} 人  --->  新: {diff_total} 人  (总变动: {diff_total - org_total})")
        print("-" * 65)

        print("\n💡 操作指南：")
        print("  [回车] 确认无误，开始处理")
        print("  [序号] 修改某班人数 (输入 数字1 数字2，数字1是班级前的序号，数字2是需要修改的人数。如修改一年级2班人数为36，则输入 2 36)")
        print("  [n]    退出程序")

        cmd = input("\n👉 请输入指令: ").strip().lower()

        if cmd == 'y' or cmd == '':
            break
        elif cmd == 'n':
            print("👋 已取消操作，再见。")
            return

        # 智能解析修改指令
        # 支持 "5 45" 格式，也支持只输入 "5" 然后追问
        parts = cmd.split()
        target_idx = -1
        new_val = -1

        try:
            target_idx = int(parts[0]) - 1
            if 0 <= target_idx < total_classes:
                if len(parts) == 2:
                    new_val = int(parts[1])
                else:
                    # 人性化追问
                    key = sorted_classes[target_idx]
                    name = format_class_name(key[0], key[1], grade_map)
                    curr = targets_map[key]
                    val_str = input(f"正在修改 【{name}】 (当前 {curr})，请输入新人数: ")
                    new_val = int(val_str)

                # 执行修改
                targets_map[sorted_classes[target_idx]] = new_val
                print("✅ 修改已更新！")
                time.sleep(0.5)  # 暂停一下让用户看到提示
            else:
                print("❌ 序号超出范围，请重试。")
                time.sleep(1)
        except:
            print("❌ 指令无法识别，请输入序号数字。")
            time.sleep(1)

    # --- 5. 执行处理 ---
    print_section("正在处理数据")
    final_dfs = []
    all_logs = []
    all_changes = []

    # 获取年级列表
    sorted_grades = []
    seen = set()
    for g, c in sorted_classes:
        if g not in seen: sorted_grades.append(g); seen.add(g)

    # 进度条效果
    for i, grade in enumerate(sorted_grades):
        # 打印进度
        grade_name = grade_map.get(grade, {}).get('name', str(grade))
        sys.stdout.write(f"\r⏳ 正在计算 {grade_name} ({i + 1}/{len(sorted_grades)})...")
        sys.stdout.flush()

        grade_df = df[df['年级'] == grade]
        processed, logs, changes = process_grade_data(grade_df, targets_map, grade)
        final_dfs.extend(processed)
        all_logs.extend(logs)
        all_changes.extend(changes)
        time.sleep(0.2)  # 模拟一点计算感

    print("\n✅ 计算完成！")

    # --- 6. 智能保存 ---
    if final_dfs:
        result_df = pd.concat(final_dfs)
        change_log_df = pd.DataFrame(all_changes)

        while True:
            try:
                # 尝试删除旧文件（如果存在）
                if os.path.exists(output_file):
                    os.remove(output_file)

                # 写入新文件
                with pd.ExcelWriter(output_file) as writer:
                    result_df = result_df[df.columns]
                    result_df.to_excel(writer, sheet_name='最终名单', index=False)

                    if not change_log_df.empty:
                        change_log_df.to_excel(writer, sheet_name='变动记录', index=False)
                    else:
                        pd.DataFrame({'提示': ['本次无人员变动']}).to_excel(writer, sheet_name='变动记录', index=False)
                break  # 成功则跳出循环

            except PermissionError:
                print(f"\n❌ 保存失败！文件 '{output_file}' 正被打开。")
                input("🔴 请关闭该 Excel 文件，然后按回车键重试...")
            except Exception as e:
                print(f"\n❌ 保存时发生未知错误: {e}")
                return

        print_section("处理结果")
        print(f"🎉 成功！文件已保存至: {output_file}")
        print(f"📊 最终总人数: {len(result_df)}")
        print(f"📋 包含两个工作表：\n   1. [最终名单] - 可直接上报\n   2. [变动记录] - 查看被删除或借调的学生详情")

        # 自动打开文件夹 (可选功能，仅限Windows)
        # os.startfile('.')

    print("\n" + "=" * 30 + " 程序结束 " + "=" * 30)
    input("按回车键关闭窗口...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        # 捕获 Ctrl+C 或 停止信号
        print("\n\n👋 程序已由用户手动停止。再见！")
        time.sleep(1)
        sys.exit(0)
    except Exception as e:
        # 捕获其他未知报错，防止闪退
        print(f"\n❌ 发生未知错误: {e}")
        input("按回车键退出...")