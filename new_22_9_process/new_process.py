import xlwt
from datetime import datetime, timedelta


def process_log(date_index, time_index, file_name, date_sign, start_sign, end_sign, opreate_sign, name, date_end_sign):
    fo = open(file_name, 'r')

    # 存放要操作的日期
    date = []

    # 利用datetime.strptime函数将时间字符串格式化

    # ！注意log文件中年月日的顺序
    # datetime_start = datetime.strptime(date_sign, "%m/%d/%Y")
    # datetime_end = datetime.strptime(date_end_sign, "%m/%d/%Y")
    datetime_start = datetime.strptime(date_sign, "%d/%m/%Y")
    datetime_end = datetime.strptime(date_end_sign, "%d/%m/%Y")

    # 将起始日期和结束日期中间的日期全部加入到date数组中
    for i in range((datetime_end - datetime_start).days + 1):
        day = str(datetime_start + timedelta(days=i)).split(" ")[0]
        day = str(int(day.split("-")[2])) + '/' + str(int(day.split("-")[1])) + '/' + str(int(day.split("-")[0]))
        date.append(day)

    # 从打开的文件中读取多行
    array = fo.readlines()

    start = []   # start数组用来存储所有匹配到的开始点
    end = []     # end数组用来存储所有匹配到的结束点

    # opreate_list数组用来存储所有中间点
    opreate_list = []
    for line in range(len(array)):
        # 如果读到的行的字符串中包含有指定的日期
        if any(day in array[line] for day in date):
            # 找到开始点
            if array[line].find(start_sign) >= 0:
                print('开始点', array[line], '--', line)
                # 加入数组中
                start.append(line)
            if array[line].find(end_sign) >= 0:
                print('结束点', array[line], '--', line)
                # 加入数组中
                end.append(line)
            if array[line].find(opreate_sign) >= 0:
                # print('move点', array[line],'--',line)
                # 加入数组中
                opreate_list.append(line)

    # 用来存储经算法检测得到的符合条件的时间段
    result = []     # result设计为二维数组，每一个成员都是一个分别存储开始点和结束点的数组

    # 将开始点序列倒序排列
    reverse_start = start
    reverse_start.reverse()

    # 找出匹配的开始点和结束点
    pre = 0  # pre保存上一个结束点

    for i in range(len(end)):
        for j in range(len(reverse_start)):
            # 定位到指定位置
            if reverse_start[j] > end[i]:
                continue
            if reverse_start[j] > pre:
                # 加入到匹配对集合中
                print('匹配到对', reverse_start[j], '到', end[i])
                temp = [reverse_start[j], end[i]]
                result.append(temp)
                break

        # 更新上一个结束点
        pre = end[i]

    # 如果result为空，说明没有符合条件的时间段
    if result.__len__() == 0:
        print('无匹配项')
        return

    # 创建一个workbook 设置编码
    workbook = xlwt.Workbook(encoding='utf-8')
    num = 1

    for i in range(len(result)):
        # 指定sheetname
        # sheet_name = 'sheet' + str(num)
        sheet_name = str(array[result[i][0]]).split(" ")[time_index].replace(':', '：') + '~' + \
                     str(array[result[i][1]]).split(" ")[time_index].replace(':', '：')

        # 创建一个worksheet
        worksheet = workbook.add_sheet(sheet_name)
        num += 1  # 表单号+1

        # 写入时间
        index = 0  # index来控制写入的行的序号

        # 起始时间
        start_flag = 1

        # 根据起始点，取得对应格式的起始时间
        time_start = array[result[i][0]].split(" ")[date_index] + ' ' + array[result[i][0]].split(" ")[time_index]
        time_start = datetime.strptime(time_start, "%m/%d/%Y %H:%M:%S")

        # 将对应的起始点和结束点之间的所有时间点都写入到表格中
        for j in range(len(opreate_list)):
            if opreate_list[j] > result[i][0] and opreate_list[j] < result[i][1]:
                # 将时间转换为需要的格式
                time = array[opreate_list[j]].split(" ")[date_index] + ' ' + array[opreate_list[j]].split(" ")[time_index]
                time = datetime.strptime(time, "%m/%d/%Y %H:%M:%S")
                # 计算出间隔的分钟
                interval = (time - time_start).total_seconds() / 60
                # 将对应的时间写入第一列
                worksheet.write(index, 0, array[opreate_list[j]].split(" ")[time_index])
                # 将对应的与起始点的间隔时间（分钟）写入第二列
                worksheet.write(index, 1, interval)
                index += 1


    # 保存
    # workbook.save(name)
    return


if __name__ == '__main__':
    # 日期在每行中的索引值
    date_index = 0
    # 时间在每行中的索引值
    time_index = 2
    # 要处理的日志的文件名
    file_name = './InLabSolution-09-17.log'
    # 要搜索的日期
    date_sign = "8/3/2022"
    # 结束的日期
    date_end_sign = "9/3/2022"
    # 开始的标志
    start_sign = '103700,13100,100000'
    # 结束的标志
    end_sign = "-10000,15000,80000"
    # 移动操作的标志
    opreate_sign = 'MOVE_REL'
    # 保存的文件名字
    name = 'move_time.xls'
    process_log(date_index=date_index, time_index=time_index, file_name=file_name, date_sign=date_sign, date_end_sign=date_end_sign, start_sign=start_sign,
                end_sign=end_sign, opreate_sign=opreate_sign, name=name)
