from tkinter.messagebox import NO
import xlrd
import xlwt
from xlutils.copy import copy
from datetime import date, datetime

import codecs
import csv


# 打开随手记模板文件
read_book = xlrd.open_workbook(r"./template.xls", formatting_info=False)

# 获取所有的sheet
print("所有的工作表：", read_book.sheet_names())
r_zhichu_sheet = read_book.sheet_names()[0]

# 根据sheet索引或者名称获取sheet内容
r_zhichu_sheet = read_book.sheet_by_index(0)

# sheet1的名称、行数、列数
print("工作表名称：%s，行数：%d，列数：%d" %
      (r_zhichu_sheet.name, r_zhichu_sheet.nrows, r_zhichu_sheet.ncols))
head: list = r_zhichu_sheet.row_values(0)  # 获取第一行的表头内容
print(head)
index = head.index('交易类型')  # 获取交易类型列所在的列数

# 拷贝一份用于写
write_book = copy(read_book)
zhichu_sheet = write_book.get_sheet(0)
assert (zhichu_sheet.get_name()=="支出")

# 支付宝的账单分类与随手记账单分类的映射
# 可以根据你的账本自定义
type_dict = {
    "交通出行": ("行车交通", "公共交通"),
    "餐饮美食": ("食品酒水", "早午晚餐"),
    "其他": ("其他杂项", "其他支出"),
    "亲友代付": ("其他杂项", "亲友代付"),
    "食品酒水": ("食品酒水", "早午晚餐"),
    "充值缴费": ("交流通讯", "手机费"),
    "日用百货": ("购物消费", "家居日用"),
    "服饰装扮": ("购物消费", "衣裤鞋帽"),
    "文化休闲": ("休闲娱乐", "电影"),
    "住房物业": ("居家生活", "物管费"),
    "生活服务": ("食品酒水", "早午晚餐"),
    "医疗健康": ("医疗教育", "药品费"),
}


def get_type(alipy_type):
    ssj_type = type_dict.get(alipy_type)
    if ssj_type == None:  # 否则返回默认的
        ssj_type = (alipy_type, "")
    return ssj_type


# 打开支付宝的账单文件
file_name = "alipay1202"
with codecs.open('./{}.csv'.format(file_name),encoding="utf8") as f:
    r = 1  # 行数
    for row in csv.DictReader(f, skipinitialspace=True):
        # print(row)
        # 原始row中有很多空格，给去除一下
        new_row = {}
        for k in row:
            new_row[k.strip()] = row[k].strip()
        print("write:", new_row)
        zhichu_sheet.write(r, head.index("交易类型"), "支出")
        zhichu_sheet.write(r, head.index("日期"), new_row["交易时间"])
        ssj_type = get_type(new_row["交易分类"])
        zhichu_sheet.write(r, head.index("分类"), ssj_type[0])
        zhichu_sheet.write(r, head.index("子分类"), ssj_type[1])
        zhichu_sheet.write(r, head.index("账户1"), "支付宝")
        # zhichu_sheet.write(r,head.index("账户2"),"支出")
        zhichu_sheet.write(r, head.index("金额"), new_row["金额"])
        zhichu_sheet.write(r, head.index("商家"), new_row["交易对方"])
        zhichu_sheet.write(r, head.index("备注"), new_row["商品说明"])
        r += 1
write_book.save("{}.xls".format(file_name))
