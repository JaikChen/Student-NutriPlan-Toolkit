import pandas as pd
import xlrd
from xlutils.copy import copy
import os

def generate_inventory_files(df, template_path, output_dir):
    """
    根据采购日期拆分数据并填充到 Excel 模板中。
    template_path and output_dir can be Path objects or strings.
    """
    template_path = str(template_path)
    output_dir = str(output_dir)
    
    grouped = df.groupby('采购日期')
    target_columns = ["食材名称", "食材单位", "食材数量", "食材单价", "小计"]
    count = 0
    errors = []

    for date, group in grouped:
        try:
            # 使用 formatting_info=True 保留模板样式 (仅支持 .xls)
            rb = xlrd.open_workbook(template_path, formatting_info=True)
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
            save_path = os.path.join(output_dir, f"{date_str}.xls")

            wb.save(save_path)
            count += 1
        except Exception as e:
            errors.append((date, str(e)))
    
    return count, errors
