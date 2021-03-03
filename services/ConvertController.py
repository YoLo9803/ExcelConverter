from ulti.XlsReader import XlsReader
from services.FormatConverter import FormatConverter
from openpyxl import Workbook
from datetime import date

class ConvertController():
    def __init__(self):
        self.__xlsReader = XlsReader()
        self.__formatConverter = FormatConverter()

    def generateProcessedXlsx(self, path):
        workbook = self.__obtainWorkbookBy(path)
        if (workbook == None):
            return
        result = self.__formatConverter.process(workbook)
        try:
            result.save('訂購憑單追蹤' + self.__getCurrentDate() + '.xlsx')
            print('轉換成功！')
        except:
            print('檔案儲存失敗！請確認該檔案是否已被開啟。')

    def __obtainWorkbookBy(self, path):
        workbook = self.__xlsReader.readXls(path)
        return workbook

    def __getCurrentDate(self):
       today = date.today()
       return today.strftime('%m%d')