#!/usr/bin/env python3

"""
    :author Wang Weiwei <email>weiwei02@vip.qq.com / weiwei.wang@100credit.com</email> 
    :sine 2017/9/26
    :version 1.0
"""
import json
import xlwt


def readjsons(param):
    """从文件中逐行读取json"""
    jsons = []
    lines = open(param, "r", encoding="utf-8").readlines()
    for line in lines:
        str2json(jsons, line)
    return jsons


def str2json(jsons, line):
    """将字符串转化为json对象，并保存到jsons列表中"""
    item = json.loads(line)
    temp = []
    repleace_json_n(item, temp)
    item["job_sec"] = "".join(temp) + "\n"
    jsons.append(item)


def repleace_json_n(item, temp):
    """去掉职位详情中的换行符与多余空格"""
    for sec in item["job_sec"]:
        sec = sec.replace("\n", "").strip()
        if sec:
            temp.append(sec)


def writeExcel(jsons, param):
    """将json文件写成excel并输出"""
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("职位信息", cell_overwrite_ok=True)
    write_excel_content(jsons, sheet)
    workbook.save(param)


def write_excel_content(jsons, sheet):
    """写excel文件内容"""
    row = 0
    for item in jsons:
        column = 0
        row = write_excel_title(column, item, row, sheet)
        for key in item.keys():
            sheet.write(row, column, item[key])
            column += 1
        row += 1


def write_excel_title(column, item, row, sheet):
    """写excel文件标题"""
    if not row:
        for key in item.keys():
            sheet.write(row, column, key)
            column += 1
        row += 1
    return row


if __name__ == '__main__':
    jsons = readjsons("../boss_products-副本.json")
    writeExcel(jsons, "boss-products.xls")
