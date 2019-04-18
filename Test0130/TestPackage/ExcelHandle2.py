# -*- coding: utf8 -*-
import xlrd
import xlwt
import time

# 本代码已废弃，新代码ExcelHandle3
# 从对应口径损益分析导出excel，需要按品种分类，第一大类是债券
# 需要修改filename
# 某些列（全价市值、应计利息、待偿期、市值、折溢摊成本、市场净价收益率、折溢摊净价收益率等）为空需要填0
# 仅适用于平均成本机构，需输入口径，目前还不支持混合估值


def main():
    filename = r"C:\Users\LHF\Desktop\指定成本与FIFO损益分析\20190408165545.xls"
    # filename = r"20190403143644.xls"
    data = xlrd.open_workbook(filename)
    newBook = xlwt.Workbook()
    newSheet = newBook.add_sheet('Capital', cell_overwrite_ok = True)
    newSheet.write(0, 2, '规模')
    newSheet.write(0, 3, '收益率')
    newSheet.write(0, 4, '待偿期')
    newSheet.write(0, 5, '综合久期')
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
    # 获取各行数据
    for i in range(0, nrows):
        # if table.cell(2, 1).ctype != 0:
        #     largeType = table.cell_value(2, 1)
        #     if table.cell(3, 2).ctype != 0:
        # print('value:', table.cell_value(i, 1))
        # 获取大类的行号
        if table.cell_value(i, 1) != '':
            lar_rowindex = i
            lar_rowindex_list.append(lar_rowindex)
        # 获取小类的行号
        if table.cell_value(i, 2) != '':
            sma_rowindex = i
            sma_rowindex_list.append(sma_rowindex)

    # 初始化规模列表
    scale_list = []
    for i in range(len(sma_rowindex_list)):
        scale_list.append(0)

    # 根据列名获取列号
    maketValueCol = getColumnIndex(table, '全价市值')
    accruedInterestCol = getColumnIndex(table, '应计利息')
    marketYieldRateCol = getColumnIndex(table, '市场净价收益率%')
    pendingPeriodCol = getColumnIndex(table, '待偿期')
    maketModifiedDurationCol = getColumnIndex(table, '市场修正久期')
    modifiedDurationCol = getColumnIndex(table, '折溢摊价格修正久期')
    tradPurposeCol = getColumnIndex(table, '交易目的')
    amortCostCol = getColumnIndex(table, '折溢摊成本')
    amortYieldRateCol = getColumnIndex(table, '折溢摊净价收益率%')

    lar_rowindex = 0
    # 写入第一个大类名称
    newSheet.write(lar_rowindex + 1, 0, table.cell_value(lar_rowindex_list[lar_rowindex], 1))


    tempScaleRate = 0
    tempScalePeriod = 0
    tempScaleDuration = 0

    # 0 -> (len-2)，循环计算每一小类，只计算到倒数第2小类规模
    for sma_rowindex in range(len(sma_rowindex_list) - 1):
        # 大类已遍历完
        if lar_rowindex < len(lar_rowindex_list) - 1:
            # 若满足，则说明已进入下一个大类
            if sma_rowindex_list[sma_rowindex] > lar_rowindex_list[lar_rowindex + 1]:
                # 写入该大类名称
                newSheet.write(sma_rowindex + 2, 0, table.cell_value(lar_rowindex_list[lar_rowindex + 1], 1))
                lar_rowindex += 1
        # 计算每小类的规模
        for pointer in range(sma_rowindex_list[sma_rowindex] + 1, sma_rowindex_list[sma_rowindex + 1]):
            # 债券不考虑负值（第一大类是债券）
            if (table.cell_type(pointer, maketValueCol) == 2 and table.cell_value(pointer, maketValueCol)) > 0 or lar_rowindex > 0:
                # 盯市：全价市值；市场净价收益率
                if caliber == 0:
                    # 剔除非数值型数据
                    if table.cell_type(pointer, maketValueCol) == 2 and table.cell_type(pointer, marketYieldRateCol) == 2:
                        tempScale = table.cell_value(pointer, maketValueCol)
                        tempRate = table.cell_value(pointer, marketYieldRateCol)
                # 折溢摊：全价市值；折溢摊净价收益率
                elif caliber == 1:
                    if table.cell_type(pointer, maketValueCol) == 2 and table.cell_type(pointer, amortYieldRateCol) == 2:
                        tempScale = table.cell_value(pointer, maketValueCol)
                        tempRate = table.cell_value(pointer, amortYieldRateCol)
                # 混合估值：根据交易目的判断
                else:
                    if table.cell_value(pointer, tradPurposeCol) == 'Trading' or '为出售而持有/剩余':
                        tempScale = table.cell_value(pointer, maketValueCol) + table.cell_value(pointer,accruedInterestCol)
                        tempRate = table.cell_value(pointer, marketYieldRateCol)
                    else:
                        tempScale = table.cell_value(pointer, amortCostCol) + table.cell_value(pointer, accruedInterestCol)
                        tempRate = table.cell_value(pointer, amortYieldRateCol)
                # if tempRate == '':
                #     tempRate = 0
                # 待偿期
                if table.cell_type(pointer, pendingPeriodCol) == 2:
                    tempPeriod = table.cell_value(pointer, pendingPeriodCol)
                    tempScalePeriod += tempScale * tempPeriod
                # 只计算债券（第一个大类）久期
                if lar_rowindex == 0:
                    # 盯市：市场修正久期
                    if caliber == 0:
                        if table.cell_type(pointer, maketModifiedDurationCol) == 2:
                            tempDuration = table.cell_value(pointer, maketModifiedDurationCol)
                    # 折溢摊：折溢摊价格修正久期
                    elif caliber == 1:
                        if table.cell_type(pointer, modifiedDurationCol) == 2:
                            tempDuration = table.cell_value(pointer, modifiedDurationCol)
                    else:
                        pass
                    tempScaleDuration += tempScale * tempDuration
                tempScaleRate += tempScale * tempRate
                scale_list[sma_rowindex] += tempScale
        # 写入该小类名称
        newSheet.write(sma_rowindex + 2, 1, table.cell_value(sma_rowindex_list[sma_rowindex], 2))
        # 写入该小类规模
        newSheet.write(sma_rowindex + 2, 2, scale_list[sma_rowindex])
        try:
            # 写入收益率
            newSheet.write(sma_rowindex + 2, 3, tempScaleRate/scale_list[sma_rowindex])
            # 写入待偿期
            newSheet.write(sma_rowindex + 2, 4, tempScalePeriod / scale_list[sma_rowindex])
            if lar_rowindex == 0:
                newSheet.write(sma_rowindex + 2, 5, tempScaleDuration / scale_list[sma_rowindex])
        except ZeroDivisionError:
            newSheet.write(sma_rowindex + 2, 3, 0)

        tempScale = 0
        tempRate = 0
        tempPeriod = 0
        tempDuration = 0
        tempScaleRate = 0
        tempScalePeriod = 0
        tempScaleDuration = 0


    sma_rowindex += 1
    pointer = sma_rowindex_list[sma_rowindex] + 1
    # 计算最后一个小类规模
    while table.cell_type(pointer, maketValueCol) == 2 and table.cell_value(pointer, maketValueCol) > 0:
        scale_list[sma_rowindex] += table.cell_value(pointer, maketValueCol) + table.cell_value(pointer, accruedInterestCol)
        pointer += 1
    # 写入该小类名称
    newSheet.write(sma_rowindex + 2, 1, table.cell_value(sma_rowindex_list[sma_rowindex], 2))
    # 写入该小类规模
    newSheet.write(sma_rowindex + 2, 2, scale_list[sma_rowindex])
    # newBook.save(r"C:\Users\LHF\Desktop\指定成本与FIFO损益分析\驾驶舱.xls")


    # 若大类比小类最后一行行号大
    while lar_rowindex < len(lar_rowindex_list) - 1:
        newSheet.write(sma_rowindex + 3, 0, table.cell_value(lar_rowindex_list[lar_rowindex + 1], 1))
        lar_rowindex += 1
        sma_rowindex += 1

    timetmp = time.strftime('%Y%m%d%H%M%S', time.localtime(time.time()))
    newBook.save(r"C:\Users\LHF\Desktop\指定成本与FIFO损益分析\驾驶舱" + timetmp + ".xls")

    # print(scale_list)


# 获得columnName对应的列号
def getColumnIndex(table, columnName):
    columnIndex = 0
    for i in range(table.ncols):
        if(table.cell_value(0, i) == columnName):
            columnIndex = i
            break
    return columnIndex


if __name__ == '__main__':
    global caliber
    print('1:盯市')
    print('2:折溢摊')
    print('3:混合估值')
    caliberNum = eval(input("请输入口径（1-3）："))
    # 盯市
    if caliberNum == 1:
        caliber = 0
    # 折溢摊
    elif caliberNum == 2:
        caliber = 1
    # 混合估值
    else:
        caliber = 2
    main()
