import xlwt


newBook = xlwt.Workbook()
newSheet = newBook.add_sheet('Capital', cell_overwrite_ok = True)
newSheet.write(1, 1, '233')
newBook.save(r"C:\Users\LHF\Desktop\指定成本与FIFO损益分析\驾驶舱.xls")
