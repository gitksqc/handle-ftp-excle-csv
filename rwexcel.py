#!/bin/env python3
# -*- coding: utf-8 -*-

from collections import namedtuple
import csv
from ftplib import FTP
import ftplib
from sqlite3 import Row
import xlrd
from xlrd import xldate_as_datetime
import xlwt
import datetime


ftp = FTP('192.168.56.1', 'user', '123123')
excel = xlrd.open_workbook('sample.xls')
sheet = excel.sheets()[0]
print(sheet.nrows)
print(sheet.ncols)

# 写入xls文件
dest_book = xlwt.Workbook(encoding="utf-8", style_compression=0)
dest_sheet = dest_book.add_sheet('Sheet1', cell_overwrite_ok=True)


n = datetime.datetime.strptime(datetime.datetime.now().strftime('%Y/%m/%d'), '%Y/%m/%d')
nn = n + datetime.timedelta(days=1)
nnn = nn.strftime('%Y/%m/%d')
print(n)
print(nn)
print(nnn)
print(nnn > '2022-04')

date_format = '%Y/%m/%d'
now_date = datetime.datetime.now().strftime(date_format)

header_line = 5
# 红色单元格
yellow_style = xlwt.XFStyle()
yellow_pattern = xlwt.Pattern()
yellow_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
yellow_pattern.pattern_fore_colour = xlwt.Style.colour_map['yellow']
yellow_style.pattern = yellow_pattern

# 黄色单元格
red_style = xlwt.XFStyle()
red_pattern = xlwt.Pattern()
red_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
red_pattern.pattern_fore_colour = xlwt.Style.colour_map['red']
red_style.pattern = red_pattern

for i in range(6, sheet.nrows):
    print(sheet.row_values(6))
    # for j in range(sheet.ncols):
        # if sheet.row_values(i)[j] == '电池编号': 
    
    battery_id = sheet.row_values(i)[0]
    app_id = sheet.row_values(i)[6]
    pro_id = sheet.row_values(i)[7]
    counter_date = sheet.row_values(i)[17]
    # 电压下限
    voltage_low_limit = float(sheet.row_values(i)[16])
    # 放电容量平均值圈数
    capacity_avg_cycle = set([2, 3, 4, 5, 6, 7, 8, 9, 10, 11])
    capacity_avg = 0.0
    # 容量保持率 12圈开始,每5圈计算一次
    capacity_keep_percent = dict()
    # 容量变化率
    capacity_change_percent = dict()
    # 温度变化率
    temp_percent = dict()
    # 容量
    capacity_dict = dict()
    # 周期
    cycle_dict = dict()

    temp = 0.0
    dirs_dict = dict()
    
    if battery_id == '':
        continue
    # print(counter_date.strip('\t'))
    # print(counter_date.isDigit())
    # print(float(counter_date))
    counter_date_str = xldate_as_datetime(float(counter_date), 0).strftime(date_format)
    counter_date_str = datetime.datetime.strptime(counter_date_str, date_format)
    counter_date_tmp = counter_date_str + datetime.timedelta(days=5)
    if counter_date_tmp.strftime(date_format) < now_date:
        print(battery_id, pro_id, counter_date_str.year, app_id)
        ftp.cwd('/')
        # print(ftp.dir())
        # print('电池验证部（新数据）/' + str(counter_date_str.year) + '/' + pro_id + '/' + app_id)
        try:
            ftp.cwd('电池验证部（新数据）/' + str(counter_date_str.year) + '/' + pro_id + '/' + app_id)
        except ftplib.error_perm:
            # print('The battery path is not existed.')
            # continue
            break
        
        #  + '/' + battery_id
        battery_dirs_dict = dict()
        # 根据电池编号定为最新的目录
        battery_dirs = ftp.mlsd()
        for bd in battery_dirs:
            print(bd)
            if battery_id == bd[0].split('-')[0]:
                battery_dirs_dict[bd[1].get('modify')] = bd[0]
        # print(battery_dirs)
        # print(battery_dirs_dict.keys())
        if len(battery_dirs_dict.keys()) <= 0:
            continue
        # print([k for k in sorted(battery_dirs_dict.keys(), reverse=True)][0])
        battery_dir = battery_dirs_dict[[k for k in sorted(battery_dirs_dict.keys(), reverse=True)][0]]
        ftp.cwd(battery_dir)
        # 查看通道目录
        file_dirs = ftp.mlsd()
        for fd in file_dirs:
            dirs_dict[fd[1].get('modify')] = fd[0]
        # the latest data dir
        source_dir = dirs_dict[[k for k in sorted(dirs_dict.keys())][0]]
        ftp.cwd(source_dir)
        # 查看csv文件
        source_files = ftp.mlsd()
        for sf in source_files:
            if sf[0].endswith('csv'):
                fh = open(sf[0], 'wb')
                ftp.retrbinary('RETR ' + sf[0], fh.write)
                fh.close()
                # 读取csv
                with open(sf[0]) as sfhandler:
                    csvhandler = csv.reader(sfhandler)
                    header = namedtuple('sourcecsv', next(csvhandler))
                    
                    for row in csvhandler:
                        data_info = header(*row)
                        # 计算第2到第11圈，放电容量平均值, 电压下限是否相等
                        cycle = int(data_info.TotalCycle)
                        if cycle >= 12:
                            break
                        if cycle in capacity_avg_cycle:
                            if data_info.StepType == 'Discharge' and voltage_low_limit == round(float(data_info.Voltage), 1):
                                capacity_avg += float(data_info.Capacity)
                        
                # 计算放电容量平均值
                capacity_avg = round(capacity_avg / len(capacity_avg_cycle), 2)
                print('capacity_avg: ', capacity_avg)
                if capacity_avg <= 0:
                    continue
                
                # 计算放电保持率，第12圈开始，每5圈计算一次
                with open(sf[0]) as sfhandler:
                    csvhandler = csv.reader(sfhandler)
                    header = namedtuple('sourcecsv', next(csvhandler))
                    
                    current_count = 0
                    for row in csvhandler:
                        data_info = header(*row)
                        cycle = int(data_info.TotalCycle)
                        capacity = float(data_info.Capacity)
                        if cycle < 12:
                            continue
                        
                        if cycle == 12:
                            if data_info.StepType == 'Discharge' and voltage_low_limit == round(float(data_info.Voltage), 1):
                                percent = round(capacity / capacity_avg, 2)
                                if percent > 1.0: percent = 1.0
                                capacity_keep_percent[cycle] = percent
                                capacity_change_percent[cycle] = 0.0
                                temp_percent[cycle] = 1.0
                                temp = float(data_info.Temp)
                                # 记录容量
                                capacity_dict[cycle] = float(data_info.Capacity)
                                # 记录周期
                                cycle_dict[cycle] = cycle
                                # print('sushun ', capacity_keep_percent, temp_percent, temp)
                        else:
                            # todo 每5圈计算一次 cycle % 5 == 0
                            if data_info.StepType == 'Discharge' and (cycle-12) % 5 == 0 and voltage_low_limit == round(float(data_info.Voltage), 1):
                                percent = round(capacity / capacity_avg, 2)
                                if percent > 1.0: percent = 1.0
                                capacity_keep_percent[cycle] = percent
                                # print('aaa, ', cycle)
                                cycle_pre = cycle - 5
                                if not cycle_pre in capacity_keep_percent.keys():
                                    continue
                                # capacity_change_percent[cycle] =  capacity_keep_percent[cycle_pre]
                                # print('now...')
                                capacity_change_percent[cycle] = round((capacity_keep_percent[cycle] - capacity_keep_percent[cycle_pre]) / capacity_keep_percent[cycle_pre], 3)
                                # 记录容量
                                capacity_dict[cycle] = float(data_info.Capacity)
                                # 记录周期
                                cycle_dict[cycle] = cycle
                                if temp <= 0:
                                    continue
                                temp_percent[cycle] = round((float(data_info.Temp) - temp) / temp, 2)
                        
        # over
        ftp.cwd('/')
        print('保持率：', capacity_keep_percent, ', 温度：', temp_percent, ', 容量变化率：', capacity_change_percent)
        

        # 将原来的行数据写入xls
        for j in range(sheet.ncols):
            # 将原来的表头写入
            dest_sheet.write(5, j, sheet.row_values(5)[j])
            # if i < 10:
                # print('su', i)
                # print(sheet.row_values(i)[1])
            dest_sheet.write(i, j, sheet.row_values(i)[j])
        
        # 将计算结果写入xls中
        # 初始容量
        # dest_sheet.write(header_line, sheet.ncols, '初始容量',  xlwt.easyxf('font:colour_index orange;'))
        dest_sheet.write(header_line, sheet.ncols, '初始容量')
        dest_sheet.write(i, sheet.ncols, capacity_avg)

        capcity_len = len(capacity_keep_percent.keys())
        print(capacity_keep_percent.keys())
        print(list(capacity_keep_percent.keys())[0])
        capacity_keys_list = list(capacity_keep_percent.keys())
        for k in range(capcity_len):
            # 写入表头
            dest_sheet.write(header_line, sheet.ncols+k*5+1, '容量保持率')
            dest_sheet.write(header_line, sheet.ncols+k*5+2, '温度变化率')
            dest_sheet.write(header_line, sheet.ncols+k*5+3, '容量变化率')
            dest_sheet.write(header_line, sheet.ncols+k*5+4, '容量')
            dest_sheet.write(header_line, sheet.ncols+k*5+5, '周期')
            # print(capacity_keep_percent.keys()[k])
            # 容量保持率
            if capacity_keep_percent[capacity_keys_list[k]] <= 0.8:
                dest_sheet.write(i, sheet.ncols+k*5+1, capacity_keep_percent[capacity_keys_list[k]], style=red_style)
            elif capacity_keep_percent[capacity_keys_list[k]] <= 0.82 and capacity_keep_percent[capacity_keys_list[k]] > 0.8:
                dest_sheet.write(i, sheet.ncols+k*5+1, capacity_keep_percent[capacity_keys_list[k]], style=yellow_style)
            else:
                dest_sheet.write(i, sheet.ncols+k*5+1, capacity_keep_percent[capacity_keys_list[k]])
            # 温度变化率
            if temp_percent[capacity_keys_list[k]] > 0.03:
                dest_sheet.write(i, sheet.ncols+k*5+2, temp_percent[capacity_keys_list[k]], style=red_style)
            else:
                dest_sheet.write(i, sheet.ncols+k*5+2, temp_percent[capacity_keys_list[k]])
            # 容量变化率
            if capacity_change_percent[capacity_keys_list[k]] >= 0.005:
                dest_sheet.write(i, sheet.ncols+k*5+3, capacity_change_percent[capacity_keys_list[k]], style=red_style)
            elif capacity_change_percent[capacity_keys_list[k]] < 0.005 and capacity_change_percent[capacity_keys_list[k]] > 0.002:
                dest_sheet.write(i, sheet.ncols+k*5+3, capacity_change_percent[capacity_keys_list[k]], style=yellow_style)
            else:
                dest_sheet.write(i, sheet.ncols+k*5+3, capacity_change_percent[capacity_keys_list[k]])
            # 容量
            dest_sheet.write(i, sheet.ncols+k*5+4, capacity_dict[capacity_keys_list[k]])
            # 周期
            dest_sheet.write(i, sheet.ncols+k*5+5, cycle_dict[capacity_keys_list[k]])
        dest_book.save('sample-dest.xls')



# import csv
# with open('sample.xls', 'rt', encoding='utf-8') as f:
#     reader = csv.reader(f)
#     for row in reader:
#         print(row)

# import openpyxl
# wb = openpyxl.load_workbook('sample.xls', read_only=True)
# sheet = wb['Sheet1']
# for datas in sheet.iter_rows(min_row=1, min_col=1):
#     # col_size = len(datas)
#     print(datas)