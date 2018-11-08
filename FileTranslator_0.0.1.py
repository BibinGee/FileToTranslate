from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import sys
from translate import Translator
import xlrd
import xlwt
import time
from xlutils.copy import copy

class Application(QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'File Translator'
        self.setGeometry(100,100,200,200)
        self.initGui()

    def initGui(self):
        layout = QVBoxLayout()
        
        self.Button = QPushButton('Load', self)
        self.Button.clicked.connect(self.on_click)

        layout.addWidget(self.Button)

        self.setLayout(layout)

    @pyqtSlot()
    def on_click(self):

        translator = Translator(to_lang = 'zh')
        
##        wbook = xlwt.Workbook(encoding = 'utf-8')
        try:
            rb = xlrd.open_workbook(r'D:\Files of Daniel\Reference file\UL Standard\UL217\UL217.xls')
            wbook = copy(rb)
            wsheet = wbook.get_sheet(0)
            
        except Exception as e:
            print(e)
##        wsheet = wbook.add_sheet('sheet1')
        
        options = QFileDialog.Options()
        
        options |= QFileDialog.DontUseNativeDialog
        
        fileName, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()")
        
        if fileName:
            print(fileName)

            
        try:
            workbook = xlrd.open_workbook(fileName)
            
            sheet = workbook.sheet_by_index(0)
            
            print(sheet.nrows)

            for i in range(0, sheet.nrows):
                
                print('row: ', i)
                
                print(sheet.row_values(i)[0])
                
                text = translator.translate(str(sheet.row_values(i)[0]))
                
                print(text)
                
                wsheet.write(i, 0, text)
                
                time.sleep(0.3)
                
                wbook.save(r'D:\Files of Daniel\Reference file\UL Standard\UL217\UL217.xls')
            
        except Exception as e:
            
            print(e)
            
        finally:
            
            workbook.release_resources()
            
            del workbook
            
            wbook.save(r'D:\Files of Daniel\Reference file\UL Standard\UL217\UL217.xls')

        
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Application()
    ex.show()
    sys.exit(app.exec_())
