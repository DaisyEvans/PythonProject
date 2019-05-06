#!usr/bin/python
# -*- coding: utf8 -*-
import xlrd
import xlwt
import time


# 从对应口径损益分析导出excel，需要按品种分类
# 需要修改filename、生成文件的路径、投组读取的路径
# 仅适用于指定成本机构，需输入口径
# 查看已实现损益，需注意损益界面选择的起止日期
# 生成的表格中展示为0或未列出的类别，可能是损益分析界面，该类别相应栏位为空，需自行检查
# 口径为折溢摊，当折溢摊净价修正久期为空时，应取盯市损益分析--市场修正久期，需自行计算
# 2019-04-19：添加已实现损益、修改债券规模为轧差统计、修改待偿期计算方法（排除同存活期规模）
# 2019-04-28: 修改待偿期计算方法（排除债券及非标资产负面额）、修改为无需对导出的损益分析表手工加数据处理、
# #             处理后的表中包含规模为0但损益不为0的数据、修改混合估值根据投组取值、导出excel数据保留位数处理
# #             修改口径为折溢摊时，折溢摊成本/折溢摊净价收益率为空时，取市值/市场净价收益率（混合估值处理相同）


def excel_handle():
    filename = r"C:\Users\LHF\Desktop\指定成本与FIFO损益分析\20190505143408.xls"
    data = xlrd.open_workbook(filename)
    newBook = xlwt.Workbook()
    newSheet = newBook.add_sheet('Capital', cell_overwrite_ok=True)
    newSheet2 = newBook.add_sheet('full_capital', cell_overwrite_ok=True)
    newSheet.write(0, 2, '规模')
    newSheet.write(0, 3, '收益率')
    newSheet.write(0, 4, '待偿期')
    newSheet.write(0, 5, '综合久期')
    newSheet.write(0, 6, '已实现损益')

    newSheet2.write(0, 2, '规模')
    newSheet2.write(0, 3, '收益率')
    newSheet2.write(0, 4, '待偿期')
    newSheet2.write(0, 5, '综合久期')
    newSheet2.write(0, 6, '已实现损益')

    # newSheet.write(0, 6, '平均收益率')
    # newSheet.write(0, 7, '平均待偿期')
    # newSheet.write(0, 8, '债券平均久期')

    table = data.sheet_by_index(0)
    # 获取行数
    nrows = table.nrows
    # 获取列数
    ncols = table.ncols
    # print("nrows %d, ncols %d" % (nrows, ncols))

    # 存放大类行号，如债券、回购、拆借等
    lar_rowindex_list = []
    # 存放小类行号，如国债、企业债、正回购等
    sma_rowindex_list = []
    # 存放大类名称
    lar_name_list = []
    # 存放小类名称
    sma_name_list = []
    # 存放处理后，第一个小类所属的大类名称（其余小类不存放），每个小类的名称、规模、收益率、久期
    data_list = []

    # 获取各行数据
    for i in range(0, nrows):
        # if table.cell(2, 1).ctype != 0:
        #     largeType = table.cell_value(2, 1)
        #     if table.cell(3, 2).ctype != 0:
        # print('value:', table.cell_value(i, 1))
        # 获取大类的行号、名称
        if table.cell_value(i, 1) != '':
            lar_rowindex = i
            lar_rowindex_list.append(lar_rowindex)
            lar_name_list.append(table.cell_value(lar_rowindex, 1))
        # 获取小类的行号、名称
        if table.cell_value(i, 2) != '':
            sma_rowindex = i
            sma_rowindex_list.append(sma_rowindex)
            sma_name_list.append(table.cell_value(sma_rowindex, 2))

    # 初始化规模列表
    scale_list = []
    for i in range(len(sma_rowindex_list)):
        scale_list.append(0)

    # 根据列名获取列号
    maketValueCol = getColumnIndex(table, '市值')
    accruedInterestCol = getColumnIndex(table, '应计利息')
    marketYieldRateCol = getColumnIndex(table, '市场净价收益率')
    pendingPeriodCol = getColumnIndex(table, '待偿期')
    modifiedDurationCol = getColumnIndex(table, '折溢摊净价修正久期')
    trade_portfolio_col = getColumnIndex(table, '交易投组')
    amortCostCol = getColumnIndex(table, '折溢摊成本')
    amortYieldRateCol = getColumnIndex(table, '折溢摊净价收益率%')
    realized_gainsCol = getColumnIndex(table, '已实现损益')

    lar_rowindex = 0
    tempScaleRate = 0
    tempScalePeriod = 0
    tempScaleDuration = 0
    tempScale_period = 0  # 存放每小类待偿期非空的规模之和

    portfolio_trading = get_trade_purpose()

    # 0 -> (len-1)，循环计算每一小类
    for sma_rowindex in range(len(sma_rowindex_list)):
        # if lar_rowindex < len(lar_rowindex_list):
        pointer_begin = sma_rowindex_list[sma_rowindex] + 1
        global pointer_end
        pointer_end = lar_rowindex_list[lar_rowindex + 1]
        if sma_rowindex < len(sma_name_list) - 1:
            pointer_end = sma_rowindex_list[sma_rowindex + 1]
            if sma_rowindex_list[sma_rowindex + 1] > lar_rowindex_list[lar_rowindex + 1]:
                lar_rowindex += 1
                pointer_end = lar_rowindex_list[lar_rowindex]
        # 计算每小类的规模
        for pointer in range(pointer_begin, pointer_end):
            # 债券轧差统计
            if table.cell_type(pointer, maketValueCol) == 2:
                # 盯市：市值+应计利息；市场净价收益率
                if caliber == 0:
                    # 剔除非数值型数据
                    if table.cell_type(pointer, accruedInterestCol) == 2:
                        tempScale = table.cell_value(pointer, maketValueCol) + table.cell_value(pointer,
                                                                                                accruedInterestCol)
                    if table.cell_type(pointer, marketYieldRateCol) == 2:
                        tempRate = table.cell_value(pointer, marketYieldRateCol)
                # 折溢摊：折溢摊成本+应计利息；折溢摊净价收益率
                elif caliber == 1:
                    # 折溢摊成本为空或为0时，取市值
                    if table.cell_type(pointer, amortCostCol) != 2 or (table.cell_type(pointer, amortCostCol) == 2 and table.cell_value(pointer, amortCostCol) == 0):
                        if table.cell_type(pointer, accruedInterestCol) == 2:
                            tempScale = table.cell_value(pointer, maketValueCol) + table.cell_value(pointer, accruedInterestCol)
                    else:
                        tempScale = table.cell_value(pointer, amortCostCol) + table.cell_value(pointer, accruedInterestCol)
                    # 折溢摊净价收益率为空或为0时，取市场净价收益率
                    if table.cell_type(pointer, amortYieldRateCol) != 2 or (table.cell_type(pointer, amortYieldRateCol) == 2 and table.cell_value(pointer, amortYieldRateCol) == 0):
                        if table.cell_type(pointer, marketYieldRateCol) == 2:
                            tempRate = table.cell_value(pointer, marketYieldRateCol)
                    else:
                        tempRate = table.cell_value(pointer, amortYieldRateCol)
                # 混合估值：根据交易目的判断
                else:
                    if table.cell_value(pointer, trade_portfolio_col) in portfolio_trading:
                        if table.cell_type(pointer, maketValueCol) == 2 and table.cell_type(pointer,
                                                                                            accruedInterestCol) == 2:
                            tempScale = table.cell_value(pointer, maketValueCol) + table.cell_value(pointer,
                                                                                                    accruedInterestCol)
                        if table.cell_type(pointer, marketYieldRateCol) == 2:
                            tempRate = table.cell_value(pointer, marketYieldRateCol)
                    else:
                        # 折溢摊成本为空或为0时，取市值
                        if table.cell_type(pointer, amortCostCol) != 2 or (
                                table.cell_type(pointer, amortCostCol) == 2 and table.cell_value(pointer,
                                                                                                 amortCostCol) == 0):
                            if table.cell_type(pointer, accruedInterestCol) == 2:
                                tempScale = table.cell_value(pointer, maketValueCol) + table.cell_value(pointer,
                                                                                                        accruedInterestCol)
                        else:
                            tempScale = table.cell_value(pointer, amortCostCol) + table.cell_value(pointer,
                                                                                                   accruedInterestCol)
                        if table.cell_type(pointer, amortYieldRateCol) != 2 or (
                                table.cell_type(pointer, amortYieldRateCol) == 2 and table.cell_value(pointer,
                                                                                                       amortYieldRateCol) == 0):
                            if table.cell_type(pointer, marketYieldRateCol) == 2:
                                tempRate = table.cell_value(pointer, marketYieldRateCol)
                        else:
                            tempRate = table.cell_value(pointer, amortYieldRateCol)

                # 待偿期
                if table.cell_type(pointer, pendingPeriodCol) == 2:
                    tempPeriod = table.cell_value(pointer, pendingPeriodCol)
                    # 待偿期加权平均时过滤债券及资产面额为负的规模
                    if (lar_name_list[lar_rowindex] != '债券小计' and lar_name_list[lar_rowindex] != '其他资产小计') or (table.cell_type(pointer, maketValueCol) == 2 and table.cell_value(pointer, maketValueCol) > 0):
                        tempScalePeriod += tempScale * tempPeriod
                        tempScale_period += tempScale
                # 只计算债券久期
                if lar_name_list[lar_rowindex] == '债券小计':
                    # 折溢摊净价修正久期
                    if table.cell_type(pointer, modifiedDurationCol) == 2:
                        tempDuration = table.cell_value(pointer, modifiedDurationCol)
                        tempScaleDuration += tempScale * tempDuration
                tempScaleRate += tempScale * tempRate
                scale_list[sma_rowindex] += tempScale

        # 已实现损益
        if table.cell_type(sma_rowindex_list[sma_rowindex], realized_gainsCol) == 2:
            realized_gains = table.cell_value(sma_rowindex_list[sma_rowindex], realized_gainsCol)
        else:
            realized_gains = 0

        # 将大类名称与第一个小类名称绑定
        if table.cell_value(sma_rowindex_list[sma_rowindex] - 1, 1) != '':
            if scale_list[sma_rowindex] != 0:
                if tempScale_period != 0:
                    data_list.append(
                        [table.cell_value(sma_rowindex_list[sma_rowindex] - 1, 1), sma_name_list[sma_rowindex],
                         scale_list[sma_rowindex],
                         tempScaleRate / scale_list[sma_rowindex],
                         tempScalePeriod / tempScale_period, tempScaleDuration / scale_list[sma_rowindex], realized_gains])
                else:
                    data_list.append(
                        [table.cell_value(sma_rowindex_list[sma_rowindex] - 1, 1), sma_name_list[sma_rowindex],
                         scale_list[sma_rowindex],
                         tempScaleRate / scale_list[sma_rowindex],
                         0, tempScaleDuration / scale_list[sma_rowindex],
                         realized_gains])
            else:
                data_list.append(
                    [table.cell_value(sma_rowindex_list[sma_rowindex] - 1, 1), sma_name_list[sma_rowindex],
                     scale_list[sma_rowindex], 0, 0, 0, realized_gains])
        else:
            if scale_list[sma_rowindex] != 0:
                if tempScale_period != 0:
                    data_list.append(
                        [0, sma_name_list[sma_rowindex],
                         scale_list[sma_rowindex],
                         tempScaleRate / scale_list[sma_rowindex],
                         tempScalePeriod / tempScale_period, tempScaleDuration / scale_list[sma_rowindex], realized_gains])
                else:
                    data_list.append(
                        [0, sma_name_list[sma_rowindex],
                         scale_list[sma_rowindex],
                         tempScaleRate / scale_list[sma_rowindex],
                         0, tempScaleDuration / scale_list[sma_rowindex],
                         realized_gains])
            else:
                data_list.append([0, sma_name_list[sma_rowindex],
                     scale_list[sma_rowindex], 0, 0, 0, realized_gains])

        tempScale, tempRate, tempPeriod, tempDuration = 0, 0, 0, 0
        tempScaleRate, tempScalePeriod, tempScaleDuration, tempScale_period = 0, 0, 0, 0

        # 处理连续若干个无小类的大类的情况
        try:
            while (sma_rowindex < len(sma_name_list) - 1) and (
                    sma_rowindex_list[sma_rowindex + 1] > lar_rowindex_list[lar_rowindex + 1]):
                lar_rowindex += 1
        except Exception:
            lar_rowindex_list.append(nrows)
            # print('小类最大行号小于大类最大行号！！')

    # print(data_list)
    i, j = 1, 1
    for data in data_list:
        if data[2] != 0 or data[6] != 0:
            # 名称、规模、收益率、待偿期、综合久期、已实现损益
            newSheet.write(i, 1, data[1])
            newSheet.write(i, 2, int(data[2]))
            newSheet.write(i, 3, '%.4f' % data[3])
            newSheet.write(i, 4, '%.4f' % data[4])
            newSheet.write(i, 5, '%.4f' % data[5])
            newSheet.write(i, 6, int(data[6]))
            if data[0] != 0:
                # 大类名称
                newSheet.write(i, 0, data[0])
            i += 1
        newSheet2.write(j, 1, data[1])
        newSheet2.write(j, 2, int(data[2]))
        newSheet2.write(j, 3, '%.4f' % data[3])
        newSheet2.write(j, 4, '%.4f' % data[4])
        newSheet2.write(j, 5, '%.4f' % data[5])
        newSheet2.write(j, 6, int(data[6]))
        if data[0] != 0:
            newSheet2.write(j, 0, data[0])
        j += 1

    timetmp = time.strftime('%Y%m%d%H%M%S', time.localtime(time.time()))
    newBook.save(r"C:\Users\LHF\Desktop\指定成本与FIFO损益分析\驾驶舱" + timetmp + ".xls")


# 获得columnName对应的列号
def getColumnIndex(table, columnName):
    columnIndex = 0
    for i in range(table.ncols):
        if table.cell_value(0, i) == columnName:
            columnIndex = i
            break
    return columnIndex


# 获取Trading类投组列表
def get_trade_purpose():
    filename = r"C:\Users\LHF\Desktop\指定成本与FIFO损益分析\投资组合维护_20190429100204.xls"
    portfolio_data = xlrd.open_workbook(filename)
    portfolio_table = portfolio_data.sheet_by_index(0)
    portfolio_rows = portfolio_table.nrows
    trading_list = ['Trading', 'FVTPL', 'Hedge', '为出售而持有/剩余']
    portfolio_trading = []
    # portfolio_not_trading = []
    for i in range(2, portfolio_rows):
        if portfolio_table.cell_value(i, 2) in trading_list :
            portfolio_trading.append(portfolio_table.cell_value(i, 1))
        # else:
        #     portfolio_not_trading.append(portfolio_table.cell_value(i, 1))
    return portfolio_trading


if __name__ == '__main__':
    global caliber
    print('1:盯市')
    print('2:折溢摊')
    print('3:混合估值')
    try:
        caliberNum = eval(input("请输入口径（1-3）："))
    except Exception:
        print('输入有误,请检查！')
    # 盯市
    if caliberNum == 1:
        caliber = 0
    # 折溢摊
    elif caliberNum == 2:
        caliber = 1
    # 混合估值
    elif caliberNum == 3:
        caliber = 2
    else:
        print('输入有误,请检查！')
    if caliber == 0 or caliber == 1 or caliber == 2:
        excel_handle()
