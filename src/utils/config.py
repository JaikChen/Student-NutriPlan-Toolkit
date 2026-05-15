import os
from pathlib import Path

# 项目根目录 (src/utils/config.py -> src/utils -> src -> root)
ROOT_DIR = Path(__file__).resolve().parent.parent.parent

# 数据目录
DATA_DIR = ROOT_DIR / 'data'
STUDENT_DATA_DIR = DATA_DIR / '1_学生名单管理'
INVENTORY_DATA_DIR = DATA_DIR / '2_食材入库管理'

# 学生名单相关路径
STUDENT_INPUT_FILE = STUDENT_DATA_DIR / '营养餐基本名单.xlsx'
STUDENT_OUTPUT_FILE = STUDENT_DATA_DIR / '营养餐_最终核定表.xlsx'
STUDENT_ARCHIVE_DIR = STUDENT_DATA_DIR / '历史备份'

# 食材入库相关路径
INVENTORY_INPUT_FILE = INVENTORY_DATA_DIR / '采购清单.xlsx'
INVENTORY_TEMPLATE_FILE = INVENTORY_DATA_DIR / '食材入库信息表.xls'
INVENTORY_OUTPUT_DIR = INVENTORY_DATA_DIR / '输出结果'
INVENTORY_ARCHIVE_DIR = INVENTORY_DATA_DIR / '历史备份'

# 自动化配置
CHROME_PROFILE_DIR = ROOT_DIR / 'chrome_profile'
TARGET_URL = "https://yyjh.xszz.edu.cn/"

def ensure_dirs():
    """确保所有必要的目录都存在"""
    dirs = [
        STUDENT_DATA_DIR,
        STUDENT_ARCHIVE_DIR,
        INVENTORY_DATA_DIR,
        INVENTORY_OUTPUT_DIR,
        INVENTORY_ARCHIVE_DIR,
        CHROME_PROFILE_DIR
    ]
    for d in dirs:
        d.mkdir(parents=True, exist_ok=True)
