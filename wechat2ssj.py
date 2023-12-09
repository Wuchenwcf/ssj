from tkinter.messagebox import NO
import xlrd
import xlwt
from xlutils.copy import copy
from datetime import date, datetime
import chardet
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


# 交易对方与随手记账单分类的映射
# 可以根据你的账本自定义
type_dict = {
    "新九天": ("行车交通", "加油"),
    "饿了么": ("食品酒水", "早午晚餐"),
    "餐饮": ("食品酒水", "早午晚餐"),
    "杭州绿烽农业有限公司": ("食品酒水", "买菜"),
    "超市": ("食品酒水", "早午晚餐"),
    "便利店": ("食品酒水", "早午晚餐"),
    "十足": ("食品酒水", "早午晚餐"),
    "杭州青青果园": ("食品酒水", "水果"),
    "手机充值": ("交流通讯", "手机费"),
    "重庆小面": ("食品酒水", "早午晚餐"),
    "众粮餐饮": ("食品酒水", "早午晚餐"),
    "中铁网络": ("行车交通", "火车"),
    "医院": ("医疗教育", "药品费"),
    "高德打车": ("行车交通", "打车"),

}

# 打开微信的账单文件
file_name = "wechat1202"
with codecs.open('./{}.csv'.format(file_name), encoding="utf-8") as f:
    r = 1  # 行数
    for row in csv.DictReader(f, skipinitialspace=True):
        # print(row)
        # 原始row中有很多空格，给去除一下
        new_row = {}
        for k in row:
            new_row[k.strip()] = row[k].strip()
        print("to write:", new_row)
        # 只记支出
        if new_row["收/支"] == "收入":
            continue
        # 退款的不用记
        if new_row["当前状态"] == "已全额退款":
            continue
        # 转账不记
        if new_row["交易类型"] == "转账":
            continue

        zhichu_sheet.write(r, head.index("交易类型"), "支出")
        zhichu_sheet.write(r, head.index("日期"), new_row["交易时间"])

        main_class = "其他杂项"
        sub_class = "其他杂项"
        for key in type_dict:
            #print(key, new_row["交易对方"], key in new_row["交易对方"])
            if key in new_row["交易对方"]:
                main_class = type_dict[key][0]
                sub_class = type_dict[key][1]
        
        # 微信的分类都比较杂，这里随便写写，导入后要手动分类
        zhichu_sheet.write(r, head.index("分类"), main_class)
        zhichu_sheet.write(r, head.index("子分类"), sub_class)
        zhichu_sheet.write(r, head.index("账户1"), "微信钱包")
        # zhichu_sheet.write(r,head.index("账户2"),"支出")
        zhichu_sheet.write(r, head.index("金额"), new_row["金额(元)"][1:])
        zhichu_sheet.write(r, head.index("商家"), new_row["交易对方"])
        zhichu_sheet.write(r, head.index("备注"), new_row["商品"])
        r += 1
write_book.save("{}.xls".format(file_name))
