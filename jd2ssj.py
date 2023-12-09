from tkinter.messagebox import NO
import xlrd
import xlwt
from xlutils.copy import copy
from datetime import date, datetime
import chardet
import codecs
import csv
from common import remove_lines_range


# 打开随手记模板文件
read_book = xlrd.open_workbook(r"./template.xls", formatting_info=False)

# 获取所有的sheet
print("所有的工作表：", read_book.sheet_names())
r_zhichu_sheet = read_book.sheet_names()[0]

# 根据sheet索引或者名称获取sheet内容
r_zhichu_sheet = read_book.sheet_by_index(0)

# sheet1的名称、行数、列数
print("工作表名称：%s,行数:%d,列数:%d" %
      (r_zhichu_sheet.name, r_zhichu_sheet.nrows, r_zhichu_sheet.ncols))
head: list = r_zhichu_sheet.row_values(0)  # 获取第一行的表头内容
print(head)
index = head.index('交易类型')  # 获取交易类型列所在的列数

# 拷贝一份用于写
write_book = copy(read_book)
zhichu_sheet = write_book.get_sheet(0)



# 打开微信的账单文件
file_name = "jd1209"
csv_file_name = "jd1209.csv"



remove_lines_range(csv_file_name, 0, 21) #去掉前20行
with codecs.open('./{}'.format(csv_file_name), encoding="utf-8") as f:
    r = 1  # 行数
    for row in csv.DictReader(f, skipinitialspace=True):
        print(row)

        # 原始row中有很多空格，给去除一下
        new_row = {}
        for k in row:
            new_row[k.strip()] = row[k].strip()

        print("to write:", new_row)
        



        zhichu_sheet.write(r, head.index("交易类型"), "支出")
        zhichu_sheet.write(r, head.index("日期"), new_row["交易时间"])

        main_class = "其他杂项"
        sub_class = "其他杂项"
        
        # 京东的分类都比较杂，这里随便写写，导入后要手动分类
        zhichu_sheet.write(r, head.index("分类"), main_class)
        zhichu_sheet.write(r, head.index("子分类"), sub_class)
        zhichu_sheet.write(r, head.index("账户1"), "京东白条")
        
        zhichu_sheet.write(r, head.index("金额"), new_row["金额"])
        zhichu_sheet.write(r, head.index("商家"), new_row["商户名称"])
        zhichu_sheet.write(r, head.index("备注"), new_row["交易说明"])

        r=r+1

write_book.save("{}.xls".format(file_name))
