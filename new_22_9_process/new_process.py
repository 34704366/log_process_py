import xlwt
from datetime import datetime, timedelta


def process_log(file_name, date_start_sign, start_sign, opreate_sign, name, date_end_sign):
    fo = open(file_name, 'r')


    # 存放要存储的日期
    date = []

    datetime_start = datetime.strptime(date_start_sign, "%m-%d")
    datetime_end = datetime.strptime(date_end_sign, "%m-%d")

    # print(datetime_start, datetime_end)

    # 将起始日期和结束日期中间的日期全部加入到date数组中
    for i in range((datetime_end - datetime_start).days + 1):
        day = str(datetime_start + timedelta(days=i)).split(" ")[0]
        day = str(day.split("-")[1]) + '-' + str(day.split("-")[2])
        date.append(day)

    # print(date)

    array = fo.readlines()

    start_points = []  # start数组用来存储所有匹配到的开始点
    end = []  # end数组用来存储所有匹配到的结束点
    # opreate_list数组用来存储所有中间点
    opreate_list = []

    result = []

    for line in range(len(array)):
        # 如果读到的行的字符串中包含有指定的日期
        if any(day in array[line] for day in date):
            # 找到开始收集的点
            if array[line].find(start_sign) >= 0:
                # print(array[line])
                # 加入数组中
                start_points.append(line)

    for i in range(len(start_points)):
        # 如果不是最后一个
        if not i == len(start_points) -1 :
            # print(start_points[i])
            for j in range(start_points[i], start_points[i+1]):
                if array[j].find(opreate_sign) >= 0:
                    result_info = []
                    # 开始采集点的信息
                    collection_info = array[start_points[i]]
                    # 移动点的信息
                    move_info = array[j]

                    result_info.append(move_info.split("|")[0])
                    result_info.append(collection_info.split("|")[1])
                    result_info.append(collection_info.split("|")[-1].split('%')[-1])

                    print(result_info)
                    result.append(result_info)
                    # 如果有多个MOVE_REL，直接跳过
                    break

    if result.__len__() == 0:
        print('没有符合条件的值')
        return

    # 创建一个workbook 设置编码
    workbook = xlwt.Workbook(encoding='utf-8')

    # 创建一个worksheet
    worksheet = workbook.add_sheet('sheet1')


    # 写入表头信息
    worksheet.write(0, 0, "该孔上移动的时间点")
    worksheet.write(0, 1, "info")
    worksheet.write(0, 2, "孔数信息")

    # 控制写入的行数
    index = 1
    # 将对应的起始点和结束点之间的所有时间点都写入到表格中
    for i in range(len(result)):
        if result.__len__() > 0:
            # 将对应的时间写入第一列
            worksheet.write(index, 0, result[i][0])
            # 将对应的时间写入第一列
            worksheet.write(index, 1, result[i][1])
            # 将对应的时间写入第一列
            worksheet.write(index, 2, result[i][2])
            index += 1

    # 保存
    workbook.save(name)
    return




if __name__ == '__main__':
    # 要处理的日志的文件名
    file_name = './InLabSolution-09-17.log'
    # file_name = './temp.log'
    # 要搜索的日期
    date_start_sign = "09-16"
    # 结束的日期
    date_end_sign = "09-17"
    # 开始收集的标志
    start_sign = 'Collect Point:'
    # 移动操作的标志
    opreate_sign = 'MOVE_REL'
    # 保存的文件名字
    name = 'move_time.xls'
    process_log(file_name=file_name, date_start_sign=date_start_sign, date_end_sign=date_end_sign, start_sign=start_sign,
                opreate_sign=opreate_sign, name=name)
