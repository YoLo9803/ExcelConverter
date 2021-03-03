from services.ConvertController import ConvertController

class UI():
    def __init__(self):
        self.__convertController = ConvertController()
    def start(self):
        while 1:
            path = self.__requireFilePath()
            if (path == 'exit'):
                break
            self.__convertController.generateProcessedXlsx(path)
    def __requireFilePath(self):
        return input('請將欲轉換檔案拉入該視窗，輸入exit關閉程式：')