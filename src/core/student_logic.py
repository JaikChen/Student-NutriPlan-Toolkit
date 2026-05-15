import re
import pandas as pd

def extract_number(text):
    """从字符串中提取第一个数字，若无则返回 0"""
    match = re.search(r'(\d+)', str(text))
    return int(match.group(1)) if match else 0

def generate_grade_map(df):
    """
    根据年级列生成排序映射表。
    支持年份格式 (如 2024) 和 普通数字格式 (如 1)。
    """
    if '年级' not in df.columns: 
        return {}
    
    unique_grades = df['年级'].dropna().unique()
    grade_data = []
    for g in unique_grades:
        num = extract_number(g)
        if num > 0: 
            grade_data.append({'raw': g, 'num': num})
    
    if not grade_data: 
        return {}
        
    max_num = max(item['num'] for item in grade_data)
    is_year_format = max_num > 1900
    base_year = max_num if is_year_format else 0
    cn_nums = {1: '一', 2: '二', 3: '三', 4: '四', 5: '五', 6: '六', 7: '七', 8: '八', 9: '九'}
    
    mapping = {}
    for item in grade_data:
        level = base_year - item['num'] + 1 if is_year_format else item['num']
        display_name = f"{cn_nums.get(level, str(level))}年级"
        mapping[item['raw']] = {'sort': level, 'name': display_name}
    return mapping

def get_class_sort_key(raw_grade, raw_class, grade_map):
    """生成班级排序键"""
    g_sort = grade_map[raw_grade]['sort'] if raw_grade in grade_map else 999
    c_num = extract_number(raw_class)
    c_sort = c_num if c_num > 0 else 999
    return (g_sort, c_sort)

def process_grade_data(grade_df, targets_map, grade_key):
    """
    核心业务逻辑：处理年级内各班级的人数调剂。
    1. 裁员：超出目标人数的学生进入待调剂池。
    2. 补员：目标人数不足的班级从未分配池中借调学生。
    """
    processed_dfs = []
    change_records = []
    spare_pool = []
    class_core_data = {}
    logs = []

    # Step 1: 收集超出人员
    for cls in grade_df['班级'].unique():
        cls_df = grade_df[grade_df['班级'] == cls]
        target = targets_map.get((grade_key, cls), len(cls_df))
        
        if len(cls_df) > target:
            class_core_data[cls] = cls_df.iloc[:target]
            spares = cls_df.iloc[target:].to_dict('records')
            for s in spares:
                s['_origin_class'] = cls
            spare_pool.extend(spares)
        else:
            class_core_data[cls] = cls_df
            
    # Step 2: 借调人员补位
    for cls, current_df in class_core_data.items():
        target = targets_map.get((grade_key, cls), len(current_df))
        needed = target - len(current_df)
        final_cls_df = current_df.copy()
        
        if needed > 0 and spare_pool:
            borrowed = []
            while needed > 0 and spare_pool:
                row = spare_pool.pop(0)
                change_records.append({
                    '年级': grade_key, 
                    '姓名': row.get('姓名', '未知'),
                    '原班级': row['_origin_class'], 
                    '操作': '借调变动',
                    '现班级': cls, 
                    '身份证号': row.get('身份证号', '')
                })
                row['班级'] = cls
                del row['_origin_class']
                borrowed.append(row)
                needed -= 1
            
            if borrowed:
                final_cls_df = pd.concat([final_cls_df, pd.DataFrame(borrowed)], ignore_index=True)
        
        processed_dfs.append(final_cls_df)

    # Step 3: 记录仍多出的人员
    for row in spare_pool:
        change_records.append({
            '年级': grade_key, 
            '姓名': row.get('姓名', '未知'),
            '原班级': row['_origin_class'], 
            '操作': '彻底删除',
            '现班级': '无', 
            '身份证号': row.get('身份证号', '')
        })
        
    return processed_dfs, logs, change_records
