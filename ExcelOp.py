import os;

from win32com.client import Dispatch


class excelOp:
    def __init__(self,xlsPath):
        self.filePath=xlsPath
       
        
    def excelChart(self):
        ex = Dispatch("Excel.Application")
        print "ex=",ex
        excelPath=self.filePath
        print excelPath
   
        wb = ex.Workbooks.open(excelPath)
        ws1=wb.WorkSheets("AAPL")
        print ws1.Cells(2,5)

        ws = wb.Worksheets.Add()
        ws.Name="New Sheet"
        #ws.Range('$A1:$D1').Value = ['NAME', 'PLACE', 'RANK', 'PRICE']
        #ws.Range('$A2:$D2').Value = ['Foo', 'Fooland', 1, 100]
        #ws.Range('$A3:$D3').Value = ['Bar', 'Barland', 2, 75]
        #ws.Range('$A4:$D4').Value = ['Stuff', 'Stuffland', 3, 50]
        #wb.SaveAs('D:\\add_a_worksheet.xlsx')
        wb.SaveAS("D:\\AAPL1.xls")
        ex.Application.Quit()
        #wb.Charts.Add()
        #wc1 = wb.Charts

    
if __name__ == "__main__":

    excelObject =excelOp("D:\\AAPL.csv")
    excelObject.excelChart();
