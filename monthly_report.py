from tkinter.tix import Tree
import pandas as pd
import openpyxl
import re
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import os
import glob
import xlwings as xw
from copy import copy  # 导入 copy 函数
import calendar
import xlrd  # 导入 xlrd 库，用于读取 .xls 文件


date_fmt = "%Y%m"
# key是tonns里面的名字，value是report中列名字
tonns_name = {}
tonns_name["25"] = "V 25"
tonns_name['24'] = "V 24"
tonns_name['27'] = "V 27"
tonns_name["32"] = "Other"
tonns_name["34"] = "Other"
tonns_name["36"] = "Other"
tonns_name["40"] = "Other"
tonns_name["24 SD"] = "V 24 SD"
tonns_name["34F"] = "V 34 (F)"
tonns_name["PA-6 recycled"] = "Second PA"
tonns_name["Caprolactam"] = "CPL"
tonns_name["PA-6"] = "PA"
tonns_name["Polyamide light-stabilized yarn 93.5 tex"] = "Yarn"
tonns_name["Polyamide not the  thermostabilized yarn 187tex"] = "Yarn"
tonns_name["yarn 144tex"] = "Yarn"
tonns_name["yarn 94tex"] = "Yarn"
tonns_name["Polyamide fiber 1 tex"] = "Yarn"
tonns_name["Polyamide fiber 0.48 tex"] = "Yarn"
tonns_name["Polyamide fiber 0.68 tex"] = "Yarn"
tonns_name["Polyamide fiber 0.33 tex"] = "Yarn"
tonns_name["NYLON 6 Dipped NTCF"] = "Tyre cord"
# key是发生额及余额表中列名称，value是tonns中列的名称
name_dict = {}
name_dict['polyamide-6(volgamid 25)'] = '25'
name_dict['polyamide-6(volgamid 27)'] = '27'
name_dict['polyamide-6(volgamid 34F)'] = '34F'
name_dict['polyamide-6(volgamid 24)'] = '24'
name_dict['polyamide-6(volgamid 24SD)'] = '24 SD'
name_dict['polyamide-6(volgamid 32)'] = '32'
name_dict['polyamide-6(volgamid 34)'] = '34'
name_dict['polyamide-6(volgamid 36)'] = '36'
name_dict['polyamide-6(volgamid 40)'] = '40'
name_dict['CAPROLACTAM'] = 'Caprolactam'
name_dict['Nylon 6 dipped NTCF'] = 'NYLON 6 Dipped NTCF'
# name_dict['在途物资-materials in transit'] = 'Inventories (tn)'
# name_dict['库存商品-Commodity Stocks'] = 'Inventories (tn)'
# name_dict['其他非主营'] = 'others '

def load_workbook_with_xlrd(path):
    if path.endswith('.xls'):
        workbook = xlrd.open_workbook(path)
        workbook_xlsx = openpyxl.Workbook()
        for sheet_name in workbook.sheet_names():
            sheet = workbook.sheet_by_name(sheet_name)
            ws = workbook_xlsx.create_sheet(title=sheet_name)
            for row_idx in range(sheet.nrows):
                for col_idx in range(sheet.ncols):
                    ws.cell(row=row_idx + 1, column=col_idx + 1, value=sheet.cell_value(row_idx, col_idx))
        return workbook_xlsx
    else:
        return load_workbook(path)
    
def convert_date_format_en(date_str):
    # 解析原始字符串到datetime对象
    date_obj = datetime.strptime(date_str, date_fmt)
    # 创建新的日期字符串，月份全写，后面跟上年份的后两位
    new_date_str = date_obj.strftime('%B %y')
    
    return new_date_str
def convert_date_format(date_str, split_char):
    # 解析原始字符串到datetime对象
    date_obj = datetime.strptime(date_str, date_fmt)
    # 获取该月的最后一天
    last_day = calendar.monthrange(date_obj.year, date_obj.month)[1]
    # 创建新的日期字符串
    new_date_str = '{:02d}{}{:02d}{}{:02d}'.format(last_day, split_char, date_obj.month, split_char, date_obj.year % 100)
    
    return new_date_str

def find_excel_files_with_keyword(folder_path, keyword):
    # 构建搜索路径，匹配所有 Excel 文件（支持 .xlsx 和 .xls）
    search_pattern = os.path.join(folder_path, f"{keyword}*.xls*")
    # 使用 glob 查找匹配的文件
    matching_files = glob.glob(search_pattern)
    
    return matching_files[0]
def save_excel_file_for_value(path):
    month_wb = load_workbook(path, data_only=True)
    month_wb.save(path)
def copy_example_sheet_add_to_monthly(monthly_path, example_sheet_path):
    monthly_wb = load_workbook(monthly_path)
    example_wb = load_workbook(example_sheet_path)
    sheet_name = monthly_wb.sheetnames[-1]
    last_month_sheet = monthly_wb[sheet_name]
    sheet_name = datetime.strptime(sheet_name, date_fmt)
    # 计算下一个月
    if sheet_name.month < 12:
        next_month_date = sheet_name.replace(month=sheet_name.month + 1)
    else:
        # 如果是12月，则增加到下一年的1月
        next_month_date = sheet_name.replace(year=sheet_name.year + 1, month=1)
    # 新sheet名称
    sheet_name = next_month_date.strftime(date_fmt)
    last_year_date = next_month_date.replace(year=next_month_date.year - 1)
    last_year_sheet = monthly_wb[last_year_date.strftime(date_fmt)]
    ex_sheet = example_wb['example']
    # 如果目标文件中已存在同名工作表，先删除
    if sheet_name in monthly_wb.sheetnames:
        monthly_wb.remove(monthly_wb[sheet_name])
    # 复制工作表到目标工作簿
    new_sheet = monthly_wb.create_sheet(sheet_name)
    # 复制单元格内容和格式
    for row in ex_sheet.iter_rows():
        for cell in row:
            target_cell = new_sheet.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                target_cell.font = copy(cell.font)
                target_cell.border = copy(cell.border)
                target_cell.fill = copy(cell.fill)
                target_cell.number_format = cell.number_format
                target_cell.protection = copy(cell.protection)
                target_cell.alignment = copy(cell.alignment)

    # 复制列宽和行高
    for col in ex_sheet.column_dimensions:
        new_sheet.column_dimensions[col] = copy(ex_sheet.column_dimensions[col])

    for row in ex_sheet.row_dimensions:
        new_sheet.row_dimensions[row] = copy(ex_sheet.row_dimensions[row])
    # 复制合并单元格
    for merged_range in ex_sheet.merged_cells.ranges:
        new_sheet.merge_cells(str(merged_range))

    return monthly_wb, new_sheet, last_month_sheet, last_year_sheet
def write_date_to_product_tbl(new_sheet, last_month_sheet, last_year_sheet):
    # stock compare
    name = convert_date_format(last_month_sheet.title, '.')
    new_sheet['C3'].value = new_sheet['C3'].value + name
    new_sheet['M3'].value = new_sheet['M3'].value + name
    name = convert_date_format(new_sheet.title, '.')
    new_sheet['I3'].value = new_sheet['I3'].value + name
    # 本月
    name = convert_date_format_en(new_sheet.title)
    new_sheet['E3'].value = new_sheet['E3'].value + name
    new_sheet['F3'].value = new_sheet['F3'].value + name
    new_sheet['R3'].value = new_sheet['R3'].value + name
    new_sheet['V3'].value = new_sheet['V3'].value + name
    # 上一月
    name = convert_date_format_en(last_month_sheet.title)
    new_sheet['O3'].value = new_sheet['O3'].value + name
    new_sheet['R4'].value = new_sheet['R4'].value + name
    new_sheet['V4'].value = new_sheet['V4'].value + name
    # 上一年
    name = convert_date_format_en(last_year_sheet.title)
    new_sheet['T3'].value = new_sheet['T3'].value + name
    # 日期
    # Total Delivery
    name = datetime.strptime(last_year_sheet.title, date_fmt)
    new_sheet['Y5'].value = '{} м {}'.format(name.month, name.year)
    # 当前日期，下面计算下一个月还会用到
    name = datetime.strptime(new_sheet.title, date_fmt)
    new_sheet['X5'].value = '{} м {}'.format(name.month, name.year)
    # TOP 10 customers
    new_sheet['B38'] = convert_date_format(new_sheet.title, '/')
    new_sheet['F38'] = convert_date_format(last_month_sheet.title, '/')
    new_sheet['N38'] = convert_date_format(last_year_sheet.title, '/')
    # 计算下一个月
    if name.month < 12:
        name = name.replace(month=name.month + 1)
    else:
        # 如果是12月，则增加到下一年的1月
        name = name.replace(year=name.year + 1, month=1)
    # 新sheet名称
    name = name.strftime(date_fmt)
    name = convert_date_format_en(name)
    new_sheet['AA3'].value = new_sheet['AA3'].value + name

    return

def copy_last_data_to_new(new_sheet, last_month_sheet, last_year_sheet):
    start_row, end_row = 7, 32
    # 复制上月数据
    col = 3 # C列
    # total 列
    for row in range(start_row, end_row + 1):
        # 获取源单元格
        source_cell = last_month_sheet.cell(row=row, column=col+6)
        # 获取目标单元格
        target_cell = new_sheet.cell(row=row, column=col, value=source_cell.value)
        target_cell.fill = copy(source_cell.fill)

    col = 4 # C列
    # Incl on the way 列
    for row in range(start_row, end_row + 1):
        # 获取源单元格
        source_cell = last_month_sheet.cell(row=row, column=col+8)
        # 获取目标单元格
        target_cell = new_sheet.cell(row=row, column=col, value=source_cell.value)
        target_cell.fill = copy(source_cell.fill)
    # 复制上月Delivery
    start_col, end_col = 6, 8  # F列到H列

    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            # 获取源单元格
            source_cell = last_month_sheet.cell(row=row, column=col)
            # 获取目标单元格
            target_cell = new_sheet.cell(row=row, column=col+9, value=source_cell.value)
            target_cell.fill = copy(source_cell.fill)
    # 复制上一年Delivery
    col = 6 # G列
    # total 列
    for row in range(start_row, end_row + 1):
        # 获取源单元格
        source_cell = last_year_sheet.cell(row=row, column=col)
        # 获取目标单元格
        target_cell = new_sheet.cell(row=row, column=col+14, value=source_cell.value)
        target_cell.fill = copy(source_cell.fill)

    col = 8 # I列
    # Incl on the way 列
    for row in range(start_row, end_row + 1):
        # 获取源单元格
        source_cell = last_year_sheet.cell(row=row, column=col)
        # 获取目标单元格
        target_cell = new_sheet.cell(row=row, column=col+13, value=source_cell.value)
        target_cell.fill = copy(source_cell.fill)
    # 复制上一年的total delivery
    col = 24 # I列
    # Incl on the way 列
    for row in range(start_row, end_row + 1):
        # 获取源单元格
        source_cell = last_year_sheet.cell(row=row, column=col)
        # 获取目标单元格
        target_cell = new_sheet.cell(row=row, column=col+4, value=source_cell.value)
        target_cell.fill = copy(source_cell.fill)
    # TOP 10 customers
    # 上个月的TOP 10 customers
    start_row = 41
    end_row = 53
    start_col = 2
    end_col = 5
    offset = 4
    for col in range(start_col, end_col + 1):
        for row in range(start_row, end_row + 1):
            # 跳过合并单元格
            if (row == 52 or row == 53) and ((col+4) == 7):
                continue
            # 获取源单元格
            source_cell = last_month_sheet.cell(row=row, column=col)
            # 获取目标单元格
            # H列是合并单元格，跳过
            if (col+offset) == 8:
                offset = 5
            target_cell = new_sheet.cell(row=row, column=col+offset, value=source_cell.value)
            target_cell.fill = copy(source_cell.fill)
    
    # 上一年的TOP 10 customers
    start_col = 2
    end_col = 3
    for col in range(start_col, end_col + 1):
        for row in range(start_row, end_row + 1):
            # 跳过合并单元格
            if (row == 52 or row == 53) and ((col+12) == 15):
                continue
            # 获取源单元格
            source_cell = last_year_sheet.cell(row=row, column=col)
            # 获取目标单元格
            target_cell = new_sheet.cell(row=row, column=col+12, value=source_cell.value)
            target_cell.fill = copy(source_cell.fill)

def copy_tonns_data_to_report(tonns_path, new_sheet):
    quantity = load_workbook(tonns_path)['quantity']
    amount = load_workbook(tonns_path)['amount']
    # 先获取purchase信息
    purchase = []
    # 解析原始字符串到datetime对象
    date_obj = datetime.strptime(new_sheet.title, date_fmt)
    col = date_obj.month + 3

    for i in range(5, 29):
        tmp = {}
        name = quantity.cell(row=i, column=2).value
        
        for key, value in tonns_name.items():
            if str(name) == key:
                exist = False
                for item in purchase:
                    if item['name'] == value:
                        tmp = item
                        exist = True
                if not exist:
                    tmp['name'] = value
                    tmp['quantity'] = 0
                    tmp['amount'] = 0
                    purchase.append(tmp)

                # quantity
                val = quantity.cell(row=i, column=col).value
                if val is not None:
                    tmp['quantity'] += float(val)
                # amount
                val = amount.cell(row=i, column=col).value
                if val is not None:
                    tmp['amount'] += float(val)/1000
                break
    
    # print(purchase)
    # 获取sales信息
    sale = []
    for i in range(32, 56):
        tmp = {}
        name = quantity.cell(row=i, column=2).value

        for key, value in tonns_name.items():
            if str(name) == key:
                exist = False
                for item in sale:
                    if item['name'] == value:
                        tmp = item
                        exist = True
                if not exist:
                    tmp['name'] = value
                    tmp['quantity'] = 0
                    tmp['amount'] = 0
                    sale.append(tmp)

                # quantity
                val = quantity.cell(row=i, column=col).value
                if val is not None:
                    tmp['quantity'] += float(val)
                # amount
                val = amount.cell(row=i, column=col).value
                if val is not None:
                    tmp['amount'] += float(val)/1000
                break

    # print(sale)
    # 填充数据到report sheet's Receipt(E列) Delivery(F列)
    start_row, end_row = 7, 32
    col_re = 5 # E列
    col_dl = 6 # F列
    for row in range(start_row, end_row + 1):
        # 获取源单元格
        name = str(new_sheet.cell(row=row, column=1).value)# 字符串匹配

        for item in purchase:
            if name == item['name']:
                tmp_amount = item['amount']
                if name == "CPL" or name == "Yarn":
                    tmp_amount = tmp_amount / new_sheet.cell(row=2, column=2).value
                mt_cell = new_sheet.cell(row=row, column=col_re, value=item['quantity'])
                qian_cell = new_sheet.cell(row=row+1, column=col_re, value=tmp_amount)
                break
        
        for item in sale:
            if name == item['name']:
                tmp_amount = item['amount']
                if name == "CPL" or name == "Yarn":
                    tmp_amount = tmp_amount / new_sheet.cell(row=2, column=2).value                
                mt_cell = new_sheet.cell(row=row, column=col_dl, value=item['quantity'])
                qian_cell = new_sheet.cell(row=row+1, column=col_dl, value=item['amount'])
                break
                
def copy_transit_data_to_report(transit_path, new_sheet):
    # 读取Excel文件
    transit = pd.read_excel(transit_path, sheet_name='在途货物余额表', header=None, engine='openpyxl')
    product_list = []
    other = {}
    other['name'] = 'Other'
    other['quantity'] = 0
    other['amount'] = 0    
    # 遍历表格，提取产品信息
    for index, value in enumerate(transit[1]):
        if index < 5: continue  

        if pd.notna(value):
            exist = False

            for key in name_dict.keys():
                if key in str(value):
                    exist = True
                    break

            if True == exist:
                tmp_dic = {}
                name = tonns_name[name_dict[key]]
                tmp_dic['name'] = name

                if pd.notna(transit.iloc[index, 14]):
                    tmp_dic['quantity'] = float(transit.iloc[index, 14])
                else:
                    tmp_dic['quantity'] = 0

                if pd.notna(transit.iloc[index, 8]):
                    tmp_dic['amount'] = float(transit.iloc[index, 15])/1000
                else:
                    tmp_dic['amount'] = 0
                # 添加到product_list
                product_list.append(tmp_dic)
            else:
                other['amount'] += float(transit.iloc[index, 15])/1000
    # 添加other
    product_list.append(other)
    # print(product_list)
    # 填充数据到report sheet's Incl on the way(L列)   
    start_row, end_row = 7, 32
    col_in = 12 # L列
    for row in range(start_row, end_row + 1):
        # 获取源单元格
        name = str(new_sheet.cell(row=row, column=1).value)# 字符串匹配

        for item in product_list:
            if name == item['name']:
                mt_cell = new_sheet.cell(row=row, column=col_in, value=item['quantity'])
                qian_cell = new_sheet.cell(row=row+1, column=col_in, value=item['amount'])
                break

def create_current_total_delivery(new_sheet, last_month_sheet):
    delivery = new_sheet
    # 当前日期，下面计算下一个月还会用到
    current_date = datetime.strptime(delivery.title, date_fmt)
    # 当年所有月份的数据的总和
    # 填充数据到report sheet's Total Delivery(X列)
    start_row, end_row = 7, 32
    col_de = 24 # X列, openpyxl, A列序号是1
    col_ct = 6 # F列

    for row in range(start_row, end_row + 1):
        if current_date.month != 1:
            delivery.cell(row=row, column=col_de).value = last_month_sheet.cell(row=row, column=col_de).value + delivery.cell(row=row, column=col_ct).value
        else:
            delivery.cell(row=row, column=col_de).value = delivery.cell(row=row, column=col_ct).value

def create_top_10_customer_table(new_sheet, last_month_sheet, last_year_sheet, receivable_path, customer_sheet):
    customer_dict = []
    receive_wb = load_workbook_with_xlrd(receivable_path)
    total_sheet = receive_wb['汇总']
    other_dict = {
        'code': '0000',
        'func_curr_balan': 0,
        'over_due': 0,
        '3month': 0,
        'simple_name': 'Other customers'
    }
    
    customer_dict.append(other_dict)

    # 获取所有客户的汇总信息,忽略最后两行
    for row in range(2, total_sheet.max_row + 1 - 2):
        customer_code = total_sheet.cell(row=row, column=1).value
        if customer_code is not None:
            func_curr_balan = total_sheet.cell(row=row, column=3).value
            func_curr_balan_bill = total_sheet.cell(row=row, column=4).value
            over_due = func_curr_balan - func_curr_balan_bill

            if len(customer_dict) < 11:
                # 如果customer_dict中少于11个客户，直接插入并排序
                tmp_item = {
                    'code': customer_code,
                    'func_curr_balan': func_curr_balan,
                    'over_due': over_due,
                    '3month': 0,
                    'simple_name': ''
                }
                customer_dict.append(tmp_item)
                customer_dict.sort(key=lambda x: x['func_curr_balan'], reverse=True)
            else:
                # 如果customer_dict中已有10个客户，比较func_curr_balan
                if func_curr_balan > customer_dict[-1]['func_curr_balan']:
                    # 如果新客户的func_curr_balan大于最后一个客户的func_curr_balan，则替换最后一个客户
                    tmp_item = {
                        'code': customer_code,
                        'func_curr_balan': func_curr_balan,
                        'over_due': over_due,
                        '3month': 0,
                        'simple_name': ''
                    }
                    customer_dict[-1] = tmp_item
                    customer_dict.sort(key=lambda x: x['func_curr_balan'], reverse=True)
                else:
                    # 否则，将func_curr_balan累加到'Other customers'
                    other_dict['func_curr_balan'] += func_curr_balan

    # 更新'Other customers'的over_due
    other_dict['over_due'] = other_dict['func_curr_balan']

    # 打印结果（可选）
    for item in customer_dict:
        print(item)
                

def main():
    print(' ')
    print('#'*20 + '开始生成report table' + '#'*20)
    # excel所在文件夹路径
    folder_path = "."  # 当前文件夹
    # 出库汇总表
    keyword = "new monthly report"
    monthly_path = find_excel_files_with_keyword(folder_path, keyword)
    # 将公式变为数值保存
    save_excel_file_for_value(monthly_path)
    # 复制模板，添加新sheet到report,返回excel，新日期sheet，相对新日期的上个月sheet，上一年sheet
    monthly_wb, new_sheet, last_month_sheet, last_year_sheet = copy_example_sheet_add_to_monthly(monthly_path, 'example.xlsx')
    # 完成product表头
    # write_date_to_product_tbl(new_sheet, last_month_sheet, last_year_sheet)
    # 复制上月和上年数据
    copy_last_data_to_new(new_sheet, last_month_sheet, last_year_sheet)
    # 复制tonns数据
    # copy_tonns_data_to_report('tonns.xlsx', new_sheet)
    # 复制在途货物余额表
    keyword = "在途货物余额表"
    # transit_path = find_excel_files_with_keyword(folder_path, keyword)
    # copy_transit_data_to_report(transit_path, new_sheet)
    # 生成total delivery数据
    # create_current_total_delivery(new_sheet, last_month_sheet)
    # 生成top 10 customer数据
    keyword = "应收账款"
    receiveable_path = find_excel_files_with_keyword(folder_path, keyword)
    customer_supply = load_workbook('供应商和客户名字.xlsx')
    customer_sheet = customer_supply['客户']
    # 读取客户信息表
    create_top_10_customer_table(new_sheet, last_month_sheet, last_year_sheet, receiveable_path, customer_sheet)

    # 保存到新的excel文件
    file_name = 'new monthly report{}.xlsx'.format(new_sheet.title)
    if os.path.exists(file_name):  # 如果文件存在
        # 删除文件
        os.remove(file_name)

    monthly_wb.active = new_sheet
    monthly_wb.save(file_name)

    print('#'*20 + '完成    report table' + '#'*20)

if __name__ == '__main__':
    main()