from typing import List, Any, Optional, Union
import os
import pandas as pd


def replace_and_add(l1: List[Any], l2: List[Any], n: int) -> List[Any]:
    """
    从n处替换l1中的值为l2
    """
    return l1[:n] + l2


def _cleck_excel_keys(df: pd.DataFrame, keys: List[str]) -> List[str]:
    """
    检查 DataFrame 中 keys 列表中不存在对应的 key
    df      : 要检查的 DataFrame
    keys    : key 列表
    """
    return [key for key in keys if key not in set(df.columns)]


def check_column_names(df: pd.DataFrame, expected_columns: List[str]):
    """
    检查 DataFrame 中 keys 列表中不存在对应的 key
    df      : 要检查的 DataFrame
    keys    : key 列表
    不存在就抛出异常
    """
    not_found = _cleck_excel_keys(df, expected_columns)
    if not_found:
        raise ValueError(f"表格中缺少以下列：{','.join(not_found)}")


def is_file_exists(file_path):
    """
    检查文件是否存在
    :param file_path: 文件的完整路径
    :return: True(如果文件存在)或False(如果文件不存在)
    """
    return os.path.exists(file_path)


def get_excel_sheets(input_file: Union[str, None] = None, sheet_index: Union[int, None] = None) -> Union[str, List[str]]:
    """
    获取传入excel表的所有sheet_name或指定位置的sheet_name
    input_file   :  传入的excel表
    sheet_index  :  指定要获取的sheet的位置,从0开始。如果不指定,返回所有的sheet_name
    """
    # 先检查文件是否存在
    if input_file is None:
        raise ValueError('没有输入要读取的excel文件')

    if not os.path.isfile(input_file):
        raise ValueError('输入的文件路径不正确')

    df = pd.read_excel(input_file, sheet_name=None)
    sheet_names = list(df.keys())
    if sheet_index is None:
        return sheet_names
    elif sheet_index < len(sheet_names):
        return sheet_names[sheet_index]
    else:
        raise ValueError('指定的 sheet_index 超出了范围')


def verify_sheet_names(sheet_names: Optional[List[str]], excel_file: str) -> List[str]:
    """
    验证 sheet_names 是否传递。如果为空，则从指定的 excel_file 文件中获取所有的 sheet_names。
    sheet_names: sheet_names
    excel_file: 从指定 excel 文件中获取所有的 sheet_names
    """
    if sheet_names is None:
        try:
            sheet_names = get_excel_sheets(excel_file)
        except Exception as e:
            print(e)
            raise RuntimeError(f"读取 {excel_file} 文件出错")
    return sheet_names
