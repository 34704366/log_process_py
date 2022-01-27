import xlwt

def process_log(file_name, date_sign, start_sign, end_sign, opreate_sign, name):

    fo = open(file_name,'r')
    # fo = open('./test.log','r')


    array = fo.readlines()
    start = []
    end = []
    date_flag = 0
    opreate_list = []
    for line in range(len(array)):
        # count += 1
        # if count < 10:
        #     print(line,end='')
        if(array[line].find(date_sign) >= 0):
            # count += 1
            # 找到开始点
            if(array[line].find(start_sign) >= 0):
                print('开始点',array[line],'--',line)
                start.append(line)
            if(array[line].find(end_sign) >= 0):
                print('结束点', array[line],'--',line)
                end.append(line)
            if array[line].find(opreate_sign) >= 0 :
                # print('move点', array[line],'--',line)
                opreate_list.append(line)

    # print(opreate_list)
    result = []
    # 将开始点序列倒序排列
    reverse_start = start
    reverse_start.reverse()
    # 找出匹配的开始点和结束点
    pre = 0   # pre保存上一个结束点
    for i in range(len(end)):

        for j in range(len(reverse_start)):
            # print(reverse_start[j])
            # 定位到指定位置
            if reverse_start[j] > end[i]:
                continue
            if reverse_start[j] > pre:
                # 加入到匹配对集合中
                print('匹配到对',reverse_start[j],'到',end[i])
                temp = [reverse_start[j], end[i]]
                result.append(temp)
                break

        # 更新上一个结束点
        pre = end[i]

    if result.__len__() == 0:
        print('无匹配项')
        return

    # 创建一个workbook 设置编码
    workbook = xlwt.Workbook(encoding = 'utf-8')
    num = 1

    for i in range(len(result)):
        # 指定sheetname
        sheet_name = 'sheet'+str(num)
        # 创建一个worksheet
        worksheet = workbook.add_sheet(sheet_name)
        num += 1      # 表单号+1
        # 写入时间
        index = 0    # index来控制写入的行的序号
        for j in range(len(opreate_list)):
            if opreate_list[j] > result[i][0] and opreate_list[j] < result[i][1]:
                worksheet.write(index, 0, array[opreate_list[j]][10:17])
                index += 1
            # print(array[opreate_list[index]][10:17])
    # 保存
    # name = 'move_time.xls'
    workbook.save(name)
    return

if __name__ == '__main__':

    # 要处理的日志的文件名
    file_name = './cyc_comp.log'
    # 要搜索的日期
    date_sign = "1/17/2022"
    # 开始的标志
    start_sign = '103700,13100,100000'
    # 结束的标志
    end_sign = "-10000,15000,80000"
    # 移动操作的标志
    opreate_sign = 'MOVE_REL'
    # 保存的文件名字
    name = 'move_time.xls'
    process_log(file_name=file_name, date_sign=date_sign, start_sign=start_sign, end_sign=end_sign, opreate_sign=opreate_sign, name=name)