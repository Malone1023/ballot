# -*- coding: utf-8 -*-

"""
"""
import sys
from PyQt5.QtWidgets import QWidget, QCheckBox, QApplication,QPushButton,QLineEdit,QLabel
from PyQt5.QtCore import Qt,QRect
from PyQt5.QtGui import QPainter, QColor, QFont
from PyQt5.QtWidgets import QDesktopWidget,QMessageBox


import xlwt,xlrd
from xlrd import open_workbook
import os,time
from xlutils.copy import copy

class Example(QWidget):
    
    def __init__(self):
        super().__init__()
        #1中签  2中奖
        self.initExcel(1)
        self.initUI()
        
        
    def initUI(self):      

        scale = 1

        width_btn     = 100 * scale
        height_btn    = 20  * scale
        btn_text_size = 10  * scale
        
        width_input   = 100 * scale
        height_input  = 20  * scale
        
        mar_bootm     = 30  * scale
        
        width_title   = 600 * scale
        height_title  = 150 * scale
        
        width_label   = 800 * scale
        height_label  = 450 * scale
        
        height_label_history  = 160 * scale
        
        title_text_size = 100 * scale
        label_text_size = 280 * scale
        label_history_size = 70 * scale
        
        

        self.m_DragPosition = self.pos()
        screen = QDesktopWidget().screenGeometry()
        self.setStyleSheet('background-color:#ffffff;')


        btn_show = QPushButton('最近十条记录', self)
        btn_show.setCheckable(True)

        btn_show.setGeometry(screen.width() - 110 * scale , screen.height() - mar_bootm - 30,width_btn,height_btn)
        btn_show.setStyleSheet("QPushButton{background-color:#cecece;border:none;color:#dddddd;font-size:12px;}"
"QPushButton:hover{background-color:#aaaaaa;}")
        btn_show.setFont(QFont("Roman times",btn_text_size,QFont.Normal))
        btn_show.clicked[bool].connect(self.onShowHistory)

        btn_change = QPushButton('切换', self)
        btn_change.setCheckable(True)
    
        btn_change.setGeometry(screen.width() - 110 * scale , screen.height() - mar_bootm,width_btn,height_btn)
        btn_change.setStyleSheet("QPushButton{background-color:#cecece;border:none;color:#dddddd;font-size:12px;}QPushButton:hover{background-color:#aaaaaa;}")
        btn_change.setFont(QFont("Roman times",btn_text_size,QFont.Normal))
        btn_change.clicked[bool].connect(self.onChange)

        self.editText = QLineEdit(self)
        self.editText.setGeometry(QRect(screen.width() - 250 * scale, screen.height() - mar_bootm , width_input , height_input))
        self.editText.setFocus()
        
        self.title = QLabel(self)
        self.title.setText(u'中签号码')
        self.title.setGeometry(0, 30 * scale,screen.width() , height_title)
        self.title.setAlignment(Qt.AlignCenter)
        self.title.setFont(QFont("Roman times",title_text_size,QFont.Bold))

        self.code = QLabel(self)
        self.code.setText('0000')
        self.code.setGeometry(0, screen.height() / 2 - height_label + 200, screen.width(), height_label)
        self.code.setAlignment(Qt.AlignCenter)
        self.code.setStyleSheet('color: red')
        self.code.setFont(QFont("Roman times",label_text_size,QFont.Bold))

        
        self.historyLabelrow1 = QLabel(self)
        self.historyLabelrow1.setGeometry(0, screen.height() / 2 - height_label + 250, screen.width(), height_label_history)
        self.historyLabelrow1.setAlignment(Qt.AlignCenter)
        self.historyLabelrow1.setStyleSheet('color: red')
        self.historyLabelrow1.setFont(QFont("Roman times",label_history_size,QFont.Bold))
        
        self.historyLabelrow2 = QLabel(self)
        self.historyLabelrow2.setGeometry(0, screen.height() / 2 - height_label + 250 + 140 * scale, screen.width(), height_label_history)
        self.historyLabelrow2.setAlignment(Qt.AlignCenter)
        self.historyLabelrow2.setStyleSheet('color: red')
        self.historyLabelrow2.setFont(QFont("Roman times",label_history_size,QFont.Bold))
        
        self.historyLabelrow1.setVisible(False)
        self.historyLabelrow2.setVisible(False)

        #字体大小需要动态设计  宽高也是需要动态设置比例
        #1920 --> 100 大致20倍
#        print(screen.width())
#        print(screen.height())

        self.showFullScreen()    #全屏显示必须放在所有组件画完以后执行
        self.show()
        
        
    def doClick(self):
        #判断是否为空，判断是否重复 QMessageBox.information(self, "提示", "号码重复，请重新输入",QMessageBox.Yes)
        self.historyLabelrow1.setVisible(False)
        self.historyLabelrow2.setVisible(False)
        self.code.setVisible(True)
        code = self.editText.text()
        if code == '':
            pass
        elif len(str(code)) > 4:
            QMessageBox.information(self, u'提示', u'请输入4位数字',QMessageBox.Yes)
            self.code.setText('')
            self.editText.setText('')
        elif self.isrepeat(code):
            QMessageBox.information(self, u'提示', u'号码重复，请重新输入',QMessageBox.Yes)
            self.code.setText('')
            self.editText.setText('')
        else:
            self.code.setText(code)
            self.save2Excel(code)
            self.editText.setText('')
            
    def onShowHistory(self):
        listData = self.getHistory()
        self.historyLabelrow1.setVisible(True)
        self.historyLabelrow2.setVisible(True)
        self.code.setVisible(False)
        row1 = ''
        row2 = ''
        i = 0
        
        for data in listData:
            data = int(data)
            print(data)
            i += 1
            if i > 5:
                row2 = str(row2) + ' ' + str(data).zfill(4)
            else:
                row1 = str(row1) + ' ' + str(data) .zfill(4)
        self.historyLabelrow1.setText(row1)
        self.historyLabelrow2.setText(row2)

    def onChange(self):
        #clean text
        self.code.setText('')
        self.historyLabelrow1.setText('')
        self.historyLabelrow2.setText('')
        if self.title.text() == u'中签号码':
            self.initExcel(2)
            self.title.setText(u'中奖号码')
        else:
            self.initExcel(1)
            self.title.setText(u'中签号码')

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_A:
            self.showFullScreen()
        if event.key() == Qt.Key_Escape:
            self.showNormal()
        if str(event.key())=='16777220':#回车
            self.doClick()
        if event.key() == Qt.Key_Enter:
            self.doClick()
        if event.key() == Qt.Key_F1:
            self.close()

    def initExcel(self,model):
        self.model = model
        path = os.path.join(os.getcwd() + '/' + self.getFileName() + '.xls')
        tody_excel = os.path.exists(path)
        # print(tody_excel)
        #判断当天excel是否存在，不存在则创建
        if not tody_excel:
            if self.model == 1:
                self.createExcel(path,u'中签号码')
            else:
                self.createExcel(path,u'中奖号码')
        else:
            pass
        

    def getFileName(self):
        #获得当前时间时间戳
        now = int(time.time())
        #转换为其他日期格式,如:"%Y-%m-%d %H:%M:%S"
        timeStruct = time.localtime(now)
        strTime = time.strftime("%Y-%m-%d", timeStruct)
        if self.model == 1:
            fileName = time.strftime("%Y-%m-%d_qian", timeStruct)
        else:
            fileName = time.strftime("%Y-%m-%d_jiang", timeStruct)
        return fileName
 
    def createExcel(self,path,name):
        workbook = xlwt.Workbook(encoding='utf-8')
        booksheet = workbook.add_sheet('Sheet 1', cell_overwrite_ok=True)
        booksheet.write(0,0,name)
        booksheet.write(0,1,0)
        workbook.save(path)

    def save2Excel(self,code):
        #判断需要写到第几行数据
        try:
            row = self.getCurrRow()
            row_target = int(row + 1)
            
            path = os.path.join(os.getcwd() + '/' + self.getFileName() + '.xls')
            workbook = open_workbook(path)
            wb = copy(workbook)                    #利用xlutils.copy下的copy函数复制
            ws = wb.get_sheet(0)
            
            ws.write(row_target, 0, str(code))
            ws.write(0,1,row_target)
            
            wb.save(path)
        except Exception as e:
            QMessageBox.information(self, u'提示', u'出现异常，请重新输入',QMessageBox.Yes)

    def getCurrRow(self):
        path = os.path.join(os.getcwd() + '/' + self.getFileName() + '.xls')
        workbook = xlrd.open_workbook(path)
        #todo 判断是否存在
        booksheet = workbook.sheet_by_index(0)
        data = booksheet.cell_value(0,1)
        return data
    
    def getHistory(self):
        list = []
        total_row = int(self.getCurrRow())
        total_row += 1
        
        path = os.path.join(os.getcwd() + '/' + self.getFileName() + '.xls')
        workbook = xlrd.open_workbook(path)
        booksheet = workbook.sheet_by_index(0)
        for row in range(total_row -1,-1,-1):
            data = 0000
            if int(total_row - row) == 11:
                break
            if int(total_row - row) == total_row :
                break
            try:
                data = booksheet.cell_value(row,0)
            except Exception as e:
                pass
            list.append(data)
        
        return list
    
    def isrepeat(self,code):
        path = os.path.join(os.getcwd() + '/' + self.getFileName() + '.xls')
        workbook = xlrd.open_workbook(path)
        booksheet = workbook.sheet_by_index(0)
        total_row = int(self.getCurrRow())
        total_row += 1
        for row in range(total_row):
            data = booksheet.cell_value(row,0)
            try:
                data = int(data)
                code = int(code)
                if data == code:
                    return True
            except Exception as e:
                pass

        return False
        
if __name__ == '__main__':
    
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())

