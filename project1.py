from xlwings import Workbook, Sheet, Range, Chart

wb = Workbook(r'C:\Users\mih\Desktop\Stock_Screener\Proj1.xlsx')

Range('A1').value = 'Two 2'

print Range('A1').value

Range('A1').value = [['Too 1', 'Foo 2', 'Foo 3'], [10.0, 20.0, 40.0]]

Range('A1').table.value  # or: Range('A1:C2').value

Sheet(1).name


chart = Chart.add(source_data=Range('A1').table)

wb.save("C:\Users\mih\Desktop\Stock_Screener\Proj1.xlsx")
