from typing import List, Any
from calendar import monthrange
import random
import pandas as pd
from .common import verify_sheet_names, check_column_names


def _generate_work_day(work_days: List[int], work_num: int, pro_num: int):
    """
        生成当前月份对应RD的工作日期
        work_days 这个月总的工作天数
        work_num 这个人这个月总的工作天数
        pro_num 这个月这个人对应的项目数
    """
    # 从 work_days 中选择 work_num 个不重复的日期
    work_day = random.sample(work_days, work_num)
    # 使用随机的权重，随机选择一个项目分配日期
    weights = [random.uniform(0.5, 1.5)
               for _ in range(pro_num)]
    projects = random.choices(
        range(pro_num), weights=weights, k=work_num)
    # 对于每个项目，先选择一个日期进行优先分配
    priority_dates = [random.choice(
        work_day) for _ in range(pro_num)]
    # 初始化输出列表
    out = [[] for _ in range(pro_num)]
    for i in range(work_num):
        day, project = work_day[i], projects[i]
        if day == priority_dates[project]:
            # 如果该日期是该项目的优先日期，则直接分配
            priority_dates[project] = None
        else:
            # 否则，使用随机的方法从可用的日期中选择一个日期进行分配
            available_dates = [
                d for d in work_day if d != priority_dates[project]]
            day = random.choice(available_dates)
        # 将日期分配给项目
        out[project].append(day)
    return out


def _generate_excel(persons: List[Any],
                    days: List[int],
                    work_days: List[int],
                    writer: pd.ExcelWriter,
                    sheet_name: List[str]):
    """
        生产一个月的表
        persons     :   当前月份处理好的数据
        days        :   当前月份的总天数(日期)
        work_days   :   当前月份工作的天数(日期)
        writer      :   通过这个在一个excel中生成多个sheet表
        sheet_name  :   当前处理的表名
        is_sort     :   是否把同一个RD的人整理到一起
    """
    table_data = []
    for person in persons:
        # 随机生成每个人每个月每个项目的工作日期
        person['rd_days'] = _generate_work_day(
            work_days, person['work_num'], person['pro_num'])
        for idx, rd in enumerate(person['rds']):
            # 构造一行数据
            row = [rd, person['name'],
                   person['work_num'], 0] + [0]*len(days)
            # 根据每天的工作情况更新一行数据
            for i, day in enumerate(work_days):
                if day in person['rd_days'][idx]:
                    if person['is_yf']:
                        row[4+i] = 8  # 研发人员就8小时
                    else:
                        row[4+i] = random.randint(1, 2)
            # 更新研发天数
            if person['is_yf']:
                row[3] = person['work_num']
            else:
                row[3] = sum(row[4:]) / 8
            table_data.append(row)
    # 构建DataFrame
    colums = ['参与项目', '姓名', '总天数', '研发天数'] + days
    df = pd.DataFrame(table_data, columns=colums)
    # 设置索引
    df = df.set_index(['参与项目'])
    df.to_excel(writer, sheet_name)


def generate_time_table(input_file: str,
                        output_file: str,
                        year: int = 0,
                        sheet_names: List[str] = None):
    """
    生成工时表
    input_file      :   生成工时表需要的源文件
    output_file     :   生成的工时表文件
    year            :   要生成的年份
    sheet_names     :   要读取的源文件的表
    """
    # 如果不传sheet_names就自动读取
    try:
        sheet_names = verify_sheet_names(sheet_names, input_file)
        with pd.ExcelWriter(output_file) as writer:
            print('开始生成工时表...')
            for yid, sheet_name in enumerate(sheet_names):
                df = pd.read_excel(
                    input_file, sheet_name=sheet_name, keep_default_na=False)
                # 先检查DataFrame中是否存在需要的key
                check_column_names(df, ['姓名', '研发/辅助', 'RD', '假日'])
                # 处理数据
                names = [s.strip() for s in df['姓名'].tolist() if s]
                is_yfs = [s.strip() for s in df['研发/辅助'].tolist() if s]
                rds = [str(s).strip() for s in df['RD'].tolist() if s]
                # 当前月份的天数
                if year != 0:
                    # 自动获取对于月份的天数
                    day_num = monthrange(year, yid+1)[1] + 1
                else:
                    # 读取表格中的天数
                    if '天数' in df.columns:
                        day_num = int(str(df['天数'][0]).strip()) + 1
                    else:
                        raise RuntimeError(f'天数 不存在，请仔细检查{sheet_name}中的列')
                # 当前月份的所有日期1～day_num
                days = list(range(1, day_num))
                # 假期日期
                jrs = [int(str(i).strip())
                       for i in df['假日'].tolist() if i != '']
                # 当前月份工作的日期
                work_days = list(set(days).difference(set(jrs)))
                # 处理RD项目的名字
                def rds_data(d): return f"RD0{d}" if d < 10 else f"RD{d}"
                # 把DataFrame中的数据处理成下面的样子
                # {
                #     'name': '张三',
                #     'work_num': 21,
                #     'rds': ['RD01', 'RD02', 'RD03'],
                #     'pro_num': 5,
                #     'is_yf':True
                # }
                datas = [{
                    'name': name,
                    'work_num': len(work_days),
                    'rds': [rds_data(int(d)) for d in r.split(",")] if "," in r else [rds_data(int(r))],
                    'pro_num': len(r.split(",")) if "," in r else 1,
                    'is_yf': is_yf == "研发"
                } for name, r, is_yf in zip(names, rds, is_yfs)]
                # 根据上面生成的数据生成对应的excel内容
                _generate_excel(datas, days, work_days, writer, sheet_name)
            print('工时表生成成功')
    except Exception as e:
        print(e)
        raise RuntimeError(f'文件{input_file}操作出错,当前的sheet_name为{sheet_name}')


def merge_time_rd(input_file: str, output_file: str) -> None:
    """
    合并工时表的同名rd
    input_file      :   工时表
    output_file     :   合并后的表的名字
    """
    try:
        sheets = verify_sheet_names(None, input_file)
        with pd.ExcelWriter(output_file) as writer:
            print('开始合并工时表的rd...')
            for sheet in sheets:
                df = pd.read_excel(input_file, sheet_name=sheet)
                # 用参与项目排序
                df = df.sort_values(by=['参与项目'])
                # 合并项目相同的行
                df = df.set_index(['参与项目', '姓名'])
                df.to_excel(writer, sheet)
            print('合并工时表的rd完成')
    except Exception as e:
        print(f"合并工时表的rd出错: {str(e)}")


def _add_sum_row(df: pd.DataFrame, data: pd.Series, index: str, row_name: str) -> pd.DataFrame:
    """
    给当前的DataFrame最后一行加上总和
    df          :   原始的DataFrame
    data        :   最后加的一行数据
    index       :   DataFrame的索引
    row_name    :   最后一行索引位置的名字
    """
    d_idx = [i+1 for i in df.index.to_list()]
    df[index] = d_idx
    # 设置索引
    df = df.set_index(index)
    # 最后新增加一行
    df.loc[len(df.index) + 1] = data
    # 获取索引列
    d_idx = df.index.to_list()
    # 修改最后的值为 row_name
    d_idx[-1] = row_name
    # 更新索引列
    df.index = pd.Series(d_idx, name=index)
    return df


def generate_wages_table(
        time_file: str,
        person_file: str,
        output_file: str,
        time_sheet_names: List[str] = None,
        person_sheet_names: List[str] = None) -> None:
    """
    生成工资表
    time_file               :   工时表
    person_file             :   研发人员工资表
    output_file             :   生成表的名字
    time_sheet_names        :   工时表中的表
    person_sheet_names      :   工时表中的表
    """
    try:
        # 读取工时表
        # 如果不传sheet_names就自动读取
        time_sheet_names = verify_sheet_names(time_sheet_names, time_file)
        # 读取研发人员工资表
        person_sheet_names = verify_sheet_names(
            person_sheet_names, person_file)
        with pd.ExcelWriter(output_file) as writer:
            print('开始生成工资表...')
            # 遍历工时表
            for idx, sheet_name in enumerate(time_sheet_names):
                # 读取对应月份工时表
                df = pd.read_excel(time_file, sheet_name=sheet_name)
                # 读取对应月份研发人员工资表
                gzb = pd.read_excel(person_file, header=1,
                                    sheet_name=person_sheet_names[idx])
                # 检查工时表中列名是否存在
                check_column_names(df, ['参与项目', '姓名', '总天数', '研发天数'])
                # 检查研发人员工资表中列名是否存在
                check_column_names(gzb, ['序号', '姓名', '月工资/元', '月研发工资/元', '月五险/元',
                                         '月研发五险/元', '总工作时间/小时', '总研发时间/小时'])
                dic = {}
                id = 0
                for _, row in df.iterrows():
                    # 去掉名字前后的空格
                    name = str(row['姓名']).strip()
                    if not dic.get(name):
                        id += 1
                        dic[name] = {
                            '序号': id,
                            '姓名': name,
                            '总工作时间/小时': row['总天数'] * 8,
                            '总研发时间/小时': row['研发天数'] * 8
                        }
                    else:
                        dic[name]['总研发时间/小时'] += row['研发天数'] * 8

                gzb_data = gzb[['姓名', '月工资/元', '月五险/元']]
                for _, row in gzb_data.iterrows():
                    # 去掉名字前后的空格
                    name = str(row['姓名']).strip()
                    if dic.get(name):
                        dic[name]['月工资/元'] = row['月工资/元']
                        dic[name]['月五险/元'] = row['月五险/元']

                df = pd.DataFrame([i for i in dic.values()])
                df['月研发工资/元'] = df['月工资/元'] * (df['总研发时间/小时'] / df['总工作时间/小时'])
                df['月研发五险/元'] = df['月五险/元'] * (df['总研发时间/小时'] / df['总工作时间/小时'])
                # 求和
                col_sum = df[['总工作时间/小时', '总研发时间/小时', '月工资/元',
                              '月五险/元', '月研发工资/元', '月研发五险/元']].sum()
                # 添加最后一行总和
                df = _add_sum_row(df, col_sum, '序号', '总和')
                df.to_excel(writer, sheet_name)
            print('工资表生成成功')
    except Exception as e:
        print(e)
        raise RuntimeError(f'操作出错,当前的sheet_name为{sheet_name}')


def generate_time_statistics(input_file: str,
                             output_file: str,
                             sheet_names: List[str] = None) -> None:
    """
    生成工时统计表
    input_file      :   输入工时表
    output_file     :   输出统计表
    sheet_names     :   要生成的sheet表名
    """
    try:
        # 如果不传sheet_names就自动读取
        sheet_names = verify_sheet_names(sheet_names, input_file)
        with pd.ExcelWriter(output_file) as writer:
            print('开始生成工时统计表...')
            for _, sheet_name in enumerate(sheet_names):
                # 读取表格数据
                df = pd.read_excel(input_file, sheet_name=sheet_name)
                # 检查工时表中列名是否存在
                check_column_names(df, ['参与项目', '姓名', '总天数', '研发天数'])
                # 获取当前月份所有的rd项目
                rds = df['参与项目'].unique()
                # 排序
                rds.sort()
                # 获取当前月份所有的人名
                names = df['姓名'].str.strip().unique()
                # 构建参数对象
                all_dict = {}
                for name in names:
                    ndf = df[df['姓名'] == name]
                    all_dict[name] = {rd: (ndf[ndf['参与项目'] == rd]['研发天数'].iloc[0] * 8.0)
                                      for rd in rds if not ndf[ndf['参与项目'] == rd].empty}
                # 生成表头
                table_title = ['序号', '姓名'] + list(rds) + ['总工时/小时']
                # 构建DataFrame
                ef = pd.DataFrame(columns=table_title)
                for idx, key in enumerate(all_dict.keys()):
                    n_data = [idx + 1, key] + \
                        [all_dict[key].get(rd, 0) for rd in rds]
                    # 计算总和
                    he = sum(n_data[2:])
                    n_data.append(he)
                    ef.loc[idx] = n_data
                # 设置序号
                ef.set_index('序号', inplace=True)
                ef.to_excel(writer, sheet_name)
            print('生成工时统计表完成')
    except Exception as e:
        print(e)
        raise RuntimeError(f'文件{input_file}操作出错,当前的sheet_name为{sheet_name}')


def generate_rd_hours(time_file: str, gz_file: str, output_file: str) -> None:
    """
    生成各研发项目工资表
    time_file       :   工时统计表
    gz_file         :   工资表
    output_file     :   生成各研发项目工资表的名字
    """
    try:
        # 获取所有月份的工时表
        time_sheets = verify_sheet_names(None, time_file)
        # 获取所有月份的工资表
        gz_sheets = verify_sheet_names(None, gz_file)
        with pd.ExcelWriter(output_file) as writer:
            print('开始生成各研发项目工资表...')
            # 生成各个月的表
            for time_sheet, gz_sheet in zip(time_sheets, gz_sheets):
                # 读取对应的sheet表
                time_df = pd.read_excel(
                    time_file, sheet_name=time_sheet, keep_default_na=False)

                gz_df = pd.read_excel(
                    gz_file, sheet_name=gz_sheet, keep_default_na=False)

                # 这个月的rd项目
                rds = time_df.columns.values.tolist()[2:-1]
                # 根据名字合并工时表和工资表
                df = pd.merge(time_df, gz_df, how='inner', on='姓名')
                # 提取需要的列
                df = df[['姓名']+rds+['月工资/元', '总工时/小时']]
                for rd in rds:
                    # 根据每个 rd的用时 / 总工时 * 月工资
                    df[rd] = (df[rd] / df['总工时/小时']) * df['月工资/元']
                # 删除掉 总工时/小时
                df = df.drop(['总工时/小时'], axis=1)
                # 求和
                last_row = df[rds + ['月工资/元']].sum()
                # 新增最后一行总和
                df.loc[len(df.index) + 1] = last_row
                df = _add_sum_row(df, last_row, '序号', '总和')
                # 修改 月工资/元 为 总工资/元
                df = df.rename(columns={'月工资/元': '总工资/元'})
                # 生成对应的表
                df.to_excel(writer, time_sheet)
            print('生成各研发项目工资表完成')
    except Exception as e:
        print(e)
        raise RuntimeError(
            f'操作出错,当前的time_sheet为{time_sheet},gz_sheet为{gz_sheet}')


def generate_rd_safe(time_file: str, gz_file: str, output_file: str) -> None:
    """
    生成各研发项目五险一金明细表
    time_file       :   工时统计表
    gz_file         :   工资表
    output_file     :   生成各研发项目五险一金表的名字
    """
    try:
        # 获取所有月份的工时表
        time_sheets = verify_sheet_names(None, time_file)
        # 获取所有月份的工资表
        gz_sheets = verify_sheet_names(None, gz_file)
        # 生成表
        with pd.ExcelWriter(output_file) as writer:
            print('开始生成各研发项目五险一金表...')
            # 生成各个月的表
            for time_sheet, gz_sheet in zip(time_sheets, gz_sheets):
                # 读取对应的sheet表
                time_df = pd.read_excel(
                    time_file, sheet_name=time_sheet, keep_default_na=False)

                gz_df = pd.read_excel(
                    gz_file, sheet_name=gz_sheet, keep_default_na=False)

                # 这个月的rd项目
                rds = time_df.columns.values.tolist()[2:-1]
                # 根据名字合并工时表和工资表
                df = pd.merge(time_df, gz_df, how='inner', on='姓名')
                # 提取需要的列
                df = df[['姓名']+rds+['月研发五险/元', '总工时/小时']]

                for rd in rds:
                    # 根据每个 rd的用时 / 总工时 * 月工资
                    df[rd] = (df[rd] / df['总工时/小时']) * df['月研发五险/元']
                # 删除掉 总工时/小时
                df = df.drop(['总工时/小时'], axis=1)
                # 求和
                last_row = df[rds + ['月研发五险/元']].sum()
                # 最后新增一行
                df = _add_sum_row(df, last_row, '序号', '总和')
                # 修改 月工资/元 为 总工资/元
                df = df.rename(columns={'月研发五险/元': '五险一金/元'})
                # 生成对应的表
                df.to_excel(writer, time_sheet)
            print('生成各研发项目五险一金表完成')
    except Exception as e:
        print(e)
        raise RuntimeError(
            f'操作出错,当前的time_sheet为{time_sheet},gz_sheet为{gz_sheet}')
