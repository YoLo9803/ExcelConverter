from openpyxl import load_workbook

class XlsReader():
    def __init__(self):
        self.__file = None

    def readXlsx(self, path):
        try:
            self.__file = load_workbook(path)
        except:
            print:('xlsx檔開啟失敗！請確認檔案是否存在。')
            return None
        return self.__file
