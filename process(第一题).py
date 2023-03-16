import xlrd
import time
import datetime
import openpyxl
import codecs
import math

def read_file(file_url):
    try:
        data = xlrd.open_workbook(file_url)
        return data
    except Exception as e:
        print(str(e))

def process_excel(workbook, file_url, save_file_url, sheet_index=0):
    """
    find temporal incontinuous data and interpolate data where the interval is less than 4 seconds.

    :param workbook:
    :param by_name: 对应的Sheet页
    :return:
    """
    table = workbook.sheet_by_index(sheet_index)  # 获得表格

    total_rows = table.nrows  # 拿到总共行数
    print("工作表名称:", table.name)
    print("行数:", total_rows)
    print("列数:", table.ncols)
    table_header = table.row_values(0)  # 表头
    print("表头", table_header)
    time_list = []
    data_list = []
    excel_list = []
    i = 0
    for one_row in range(1, total_rows):  # 也就是从Excel第二行开始，第一行表头不算
        i += 1
        row = table.row_values(one_row)
        if row:
            if time_list != []:
                endTime = row[0][:-5]
                startTime = time_list[-1][:-5]
                endTime = datetime.datetime.strptime(endTime,"%Y/%m/%d %H:%M:%S")
                startTime = datetime.datetime.strptime(startTime,"%Y/%m/%d %H:%M:%S")
                
                # 相减得到秒数
                seconds = (endTime - startTime).seconds

                if(seconds > 1):
                    data_list.append([seconds,i])

            time_list.append(row[0])
            row[0] = row[0][:-5]
            excel_list.append(row)

    f = codecs.open(save_file_url[:-5] + '_第1问_插值行数.txt', 'w', 'utf-8')
    header = u'在原文件第几行' + '\t' + u'插值缺少几行' + '\n'
    f.write(header)
    for data in data_list:
        f.write(str(data[1]) + '\t' + str(data[0] - 1) + '\n')

    f.write(str('处理前列数 \t' + str(len(excel_list)) + '\n'))

    delay = 0
    for data in data_list:
        seconds = data[0]
        index = data[1]

        if(seconds < 5):
            print('相差几秒:',seconds,'  在第',index,'行到',index + 1,'行')

            end = excel_list[index - 2 + delay][1]
            endTime = excel_list[index - 2 + delay][0]
            endTime = datetime.datetime.strptime(endTime,"%Y/%m/%d %H:%M:%S")

            begin = excel_list[index - 1 + delay][1]
            insert_list = []
            for i in range(seconds - 1):
                insert_list.append([(endTime + datetime.timedelta(seconds = 1) * (i + 1)).strftime("%Y/%m/%d %H:%M:%S"), (begin - end) / seconds * (i + 1) + end] )
            excel_list[index - 1 + delay : index - 1 + delay] = insert_list
            delay += seconds - 1


    f.write(str('处理后列数 \t' + str(len(excel_list)) + '\n'))
    f.close()

    excel_list.insert(0, table_header)

    return excel_list

def compute_accelaration(excel_list):
    excel_list[0].insert(2, '加速度')
    for index in range(len(excel_list[1:-1])):
        index += 1
        row = excel_list[index]
        next_row = excel_list[index + 1]
        endTime = row[0]
        startTime = next_row[0]
        endTime = datetime.datetime.strptime(endTime,"%Y/%m/%d %H:%M:%S")
        startTime = datetime.datetime.strptime(startTime,"%Y/%m/%d %H:%M:%S")
            
        # 相减得到秒数
        seconds = (startTime - endTime).seconds

        if seconds == 1:
            row.insert(2, round(float(next_row[1]) - float(row[1]), 2))
        else:
            row.insert(2, '')
    return excel_list

def read_list(workbook, sheet_index=0):
    table = workbook.sheet_by_index(sheet_index)  # 获得表格
    total_rows = table.nrows  # 拿到总共行数
    table_header = table.row_values(0)  # 表头

    excel_list = []

    for one_row in range(1, total_rows):  # 也就是从Excel第二行开始，第一行表头不算
        row = table.row_values(one_row)
        if row:
            excel_list.append(row)

    return excel_list

def find_quick_acceleration(excel_list):
    for index in range(len(excel_list[1:-5])):
        row = excel_list[index]
        speed = row[1]
        time = datetime.datetime.strptime(row[0],"%Y/%m/%d %H:%M:%S")
        next_row = excel_list[index + 6]
        next_speed = next_row[1]
        next_time = datetime.datetime.strptime(next_row[0],"%Y/%m/%d %H:%M:%S")

        #print(index, speed, next_speed)
        if(float(next_speed) - float(speed) >= 100.00 and (next_time - time).seconds == 6 ):
            print('find_quick_acceleration: ', index, ' time:', row[0], float(next_speed) - float(speed))



def find_speed_mutation(excel_list, save_file_url):
    f = codecs.open(save_file_url[:-5] + '_第1问_加速度突变异常情况.txt', 'w', 'utf-8')
    header = u'突变位置\n'
    f.write(header)

    speed_mutation = []
    temp = 0
    for index in range(len(excel_list[1:-6])):
        if temp > 0:
            temp -= 1
            continue

        row = excel_list[index]
        time = datetime.datetime.strptime(row[0],"%Y/%m/%d %H:%M:%S")
        if row[2] != '':
            
            acceleration = float(row[2])
            if acceleration > 10:
                for i in range(7):
                    next_row = excel_list[index + i + 1]
                    next_time = datetime.datetime.strptime(next_row[0],"%Y/%m/%d %H:%M:%S")
                    if next_row[2] != '' :
                        if (next_time - time).seconds == i + 1:
                            next_acc = float(next_row[2])
                            if math.fabs(next_acc) <= 1.0:
                                continue
                            elif next_acc < -10: 
                                speed_mutation.append(index)
                                info = '第' + str(index + 2) + '行到第' + str(index + i + 3) + '行'
                                print('突变数据在' + info)
                                f.write(str(index + 2) + '\t' + str(index + i + 3) + '\n')
                                temp = i + 1
                                break
                            else:
                                break

            elif acceleration < -10:
                flag = 0
                for i in range(7):
                    next_row = excel_list[index + i + 1]
                    next_time = datetime.datetime.strptime(next_row[0],"%Y/%m/%d %H:%M:%S")
                    if next_row[2] != '' :
                        if (next_time - time).seconds == i + 1:
                            next_acc = float(next_row[2])
                            if math.fabs(next_acc) <= 1.0:
                                continue
                            elif next_acc > 10: 
                                speed_mutation.append(index)
                                info = '第' + str(index + 2) + '行到第' + str(index + i + 3) + '行'
                                print('突变数据在' + info)
                                f.write(str(index + 2) + '\t' + str(index + i + 3) + '\n')
                                temp = i + 1
                                flag = 1
                                break
                            else:
                                break

    f.close()

    return speed_mutation

def find_abnormal_accelration(excel_list, save_file_url):
    f = codecs.open(save_file_url[:-5] + '_第1问_加速度异常情况.txt', 'w', 'utf-8')
    header = u'在原文件第几行' + '\t' + u'时间' + '\t' + u'加速度值' '\n'
    f.write(header)

    large_acceleration = []
    large_deceleration = []

    for index in range(len(excel_list[1:])):
        row = excel_list[index]
        if row[2] != '':
            acceleration = float(row[2])
            if acceleration > 14:
                print('find_abnormal_accelration: ', index + 2, ' time:', row[0],' acceleration too large', row[2])
                large_acceleration.append(index + 1)
                f.write(str(index + 2) + '\t' + str(row[0]) + '\t' + str(row[2]) +'\n')
            elif acceleration < -28.8:
                print('find_abnormal_accelration: ', index + 2, ' time:', row[0],' deceleration too large', row[2])
                large_deceleration.append(index + 1)
                f.write(str(index + 2) + '\t' + str(row[0]) + '\t' + str(row[2]) +'\n')
    
    f.close()
    return large_acceleration, large_deceleration


def write_file(excel_list,file_url):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = '原始数据1'
    for i in range(0, len(excel_list)):
        for j in range(0, len(excel_list[i])):
            sheet.cell(row=i+1, column=j+1, value=str(excel_list[i][j]))

    workbook.save(file_url)
    print('写入数据成功')

if __name__ == '__main__':
    file_url='./处理数据/文件3.xlsx'#'./原始数据/文件3.xlsx'
    #save_file_url='./处理数据/文件1.xlsx'
    workbook = read_file(file_url)
    excel_list = read_list(workbook)
    #excel_list = process_excel(workbook, file_url, save_file_url)
    #write_file(excel_list, save_file_url)  # 输出的数据中'时间'去掉了最后的.000.
    #excel_list = compute_accelaration(excel_list)
    #write_file(excel_list, save_file_url)
    #find_quick_acceleration(excel_list)
    #large_acceleration, large_deceleration = find_abnormal_accelration(excel_list, file_url)
    speed_mutation = find_speed_mutation(excel_list, file_url)

