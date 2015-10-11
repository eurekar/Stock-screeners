from xlwings import Workbook, Sheet, Range, Chart

wb_1= Workbook(r'C:\Users\mih\Desktop\Stock_Screener\fidelity_long_screener.xls')

Stock_list=Range('Sheet1','A2').vertical.value

print Stock_list

my_file=open("IBD_import.txt","w")

for i in Stock_list:
    my_file.write(i+' ') 

my_file.close()
