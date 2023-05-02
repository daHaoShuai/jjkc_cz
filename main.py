import os
from core import generate_time_table, merge_time_rd, generate_wages_table,\
    generate_time_statistics, generate_rd_hours, generate_rd_safe
from core.common import is_file_exists

ORIGINAL_TABLE = i if (i := input('输入加计扣除工时生成原始表:')
                       ) else 'inputs/加计扣除工时生成原始表.xlsx'

YFGZ_TABLE = i if (i := input('输入研发人员工资表:')) else 'inputs/研发人员工资表.xlsx'
YEAR = i if (i := input('生成的年份:')) else 0
TIME_TABLE = 'outputs/工时表.xlsx'
TIME_TABLE1 = 'outputs/工时表(合并RD).xlsx'
GZ_TABLE = 'outputs/工资表.xlsx'
TIME_TJ_TABLE = 'outputs/工时统计表.xlsx'
YFXMGZ_TABLE = 'outputs/研发项目工资表.xlsx'
WXYJTJ_TABLE = 'outputs/各研发项目五险一金明细表.xlsx'


def init():
    if not is_file_exists(ORIGINAL_TABLE):
        raise IOError(f'文件不存在{ORIGINAL_TABLE}')
    if not is_file_exists(YFGZ_TABLE):
        raise IOError(f'文件不存在{YFGZ_TABLE}')
    if not is_file_exists('outputs'):
        os.makedirs('outputs')


if __name__ == '__main__':
    init()
    # 生成工时表
    generate_time_table(ORIGINAL_TABLE, TIME_TABLE, YEAR)
    # 合并工时表同名字RD
    merge_time_rd(TIME_TABLE, TIME_TABLE1)
    # 生成工资表
    generate_wages_table(TIME_TABLE, YFGZ_TABLE, GZ_TABLE)
    # 生成工时统计表
    generate_time_statistics(TIME_TABLE, TIME_TJ_TABLE)
    # 生成研发项目工资表
    generate_rd_hours(TIME_TJ_TABLE, GZ_TABLE, YFGZ_TABLE)
    # 生成各研发项目五险一金明细表
    generate_rd_safe(TIME_TJ_TABLE, GZ_TABLE, WXYJTJ_TABLE)
