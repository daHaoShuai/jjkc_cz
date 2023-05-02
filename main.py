from core import generate_time_table, merge_time_rd, generate_wages_table,\
    generate_time_statistics, generate_rd_hours, generate_rd_safe


ORIGINAL_TABLE = 'inputs/加计扣除工时生成原始表.xlsx'
YFGZ_TABLE = 'inputs/研发人员工资表.xlsx'
TIME_TABLE = '工时表.xlsx'
TIME_TABLE1 = '工时表(合并RD).xlsx'
GZ_TABLE = '工资表.xlsx'
TIME_TJ_TABLE = '工时统计表.xlsx'
YFXMGZ_TABLE = '研发项目工资表.xlsx'
WXYJTJ_TABLE = '各研发项目五险一金明细表.xlsx'

# 生成工时表
generate_time_table(ORIGINAL_TABLE, TIME_TABLE)
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
