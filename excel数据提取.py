# -*- coding: utf-8 -*-
# @Time    : 2022/4/5 17:56
# @Author  : Wu Haoran
# @File    : excel数据提取.py

import xlrd
import string


class excelData:
    # 初始化文件，传入文件名， 工作表名
    def __init__(self, file_name: str, table_name: str) -> None:
        f = xlrd.open_workbook(file_name)
        table = f.sheet_by_name(table_name)
        self.col_nums = table.ncols
        self.row_nums = table.nrows
        self.__table__ = table

    # 内部私有方法，对传入的列名进行处理，转换成对应的列数
    def __col_name_preprocess__(self, col_name: str) -> int:
        col = col_name.upper()
        word_to_num = dict()
        word_list = string.ascii_uppercase
        for item in enumerate(word_list):
            num = item[0]
            word = item[1]
            word_to_num[word] = num
        return word_to_num[col]

    # 传入列的名称，输出对应的一整列的数据
    def col_select(self, col_name: str) -> list:
        col = self.__col_name_preprocess__(col_name)
        col_data = self.__table__.col_values(col)
        return col_data

    # 传入行的数字，输出对应的一整行的数据
    def row_select(self, row_name: int) -> list:
        row = row_name
        row_data = self.__table__.row_values(row)
        return row_data


# 测试数据
if __name__ == '__main__':
    file_name = "测试数据.xlsx"
    table_name = "工作表1"
    method = excelData(file_name, table_name)

    A_values = method.col_select("A")
    values1 = method.row_select(2)

    print(method.col_nums)
    print(method.row_nums)
    print(A_values)
    print(values1)
