import xlsxwriter  # 导入模块
import time
from datetime import date, timedelta
import random

workbook = xlsxwriter.Workbook('机时.xlsx')  # 新建excel表

worksheet = workbook.add_worksheet('sheet1')  # 新建sheet（sheet的名称为"sheet1"）


# data = [
#
#     ['2017-9-1', '2017-9-2', '2017-9-3', '2017-9-4', '2017-9-5', '2017-9-6'],
#
#     [10, 40, 50, 20, 10, 50],
#
#     [30, 60, 70, 50, 40, 30],
#
# ]  # 自己造的数据

date_start = date(2016, 3, 20)
date_end = date(2017, 1, 1)


date_tmp = date_start

# col_date = []
#
# while date_tmp <= date_end:
#     col_date.append(date_tmp.strftime("%Y%m%d"))
#     period = random.randint(1, 3)
#     date_tmp = date_tmp + timedelta(days=period)
#     # print(col_date)
#
# row_num = len(col_date)
# todo：机时数
times = [10, 24, 24]

col_run_time = []
# while len(run_time) < row_num:

total_run_time = 0
# 总机时数 total_run_time
while total_run_time < 1800:
    time_tmp = random.choice(times)
    col_run_time.append(time_tmp)
    total_run_time += time_tmp

row_num = len(col_run_time)

col_date = []
while len(col_date) < row_num and date_tmp < date_end:
    col_date.append(int(date_tmp.strftime("%Y%m%d")))
    # TODO:日期跨度
    period = random.randint(1, 4)
    date_tmp = date_tmp + timedelta(days=period)
    # print(col_date)


# run_time = [10] * row_num

# names = ['王立志', '裴恬', '刘航', '叶海钧', '徐世琨', '张哲']
name_exp = {"王立志": ["配电主站调试", 15505295815],
            "裴恬": ["配网通信安全仿真", 15505295815],
            "刘航": ["新能源控制", 15505295815],
            "叶海钧": ["新能源控制", 155],
            "徐世琨": ["配电主站调试", 155],
            "张哲": ["配网通信安全仿真", 155],
            "付灿宇": ["DSP控制单元调试", 155]}
col_name = []
col_exp = []
col_phone_num = []

while len(col_name) < row_num:
    # name = random.sample(names, 1)
    k = random.choice(list(name_exp.keys()))
    exp = name_exp[k][0]
    phone_num = name_exp[k][1]
    col_name.append(k)
    col_exp.append(exp)
    col_phone_num.append(phone_num)

headings = ['日期', '机时数', '使用人员类型', '使用人员', '实验名称', '实验编号', '手机号码']  # 设置表头

candidate = ['校内'] * row_num
num_exp = [16017832] * row_num

worksheet.write_row('A1', headings)
worksheet.write_column('A2', col_date)
worksheet.write_column('B2', col_run_time)
worksheet.write_column('C2', candidate)
worksheet.write_column('D2', col_name)
worksheet.write_column('E2', col_exp)
worksheet.write_column('F2', num_exp)
worksheet.write_column('G2', col_phone_num)




# worksheet.write_row('A1', headings)
#
# worksheet.write_column('A2', data[0])
#
# worksheet.write_column('B2', data[1])
#
# worksheet.write_column('C2', data[2])  # 将数据插入到表格中

workbook.close()  # 将excel文件保存关闭，如果没有这一行运行代码会报错