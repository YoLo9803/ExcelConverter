from xlrd import open_workbook_xls

class XlsReader():
    def __init__(self):
        self.__file = None

    def readXls(self, path):
        try:
            self.__file = open_workbook_xls(path)
        except:
            print('檔案開啟失敗！請確認檔案是否存在。')
            return None
        return self.__file
