# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'mgrsconverter2.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QMessageBox
import mgrs
import pandas as pd

class Ui_MainWindow(QtWidgets.QMainWindow):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(385, 467)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(40, 40, 321, 51))
        font = QtGui.QFont()
        font.setFamily("Consolas")
        font.setPointSize(30)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(120, 390, 151, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(240, 130, 111, 41))
        self.pushButton.clicked.connect(self.fileUpload)


        font = QtGui.QFont()
        font.setFamily("Agency FB")
        font.setBold(True)
        font.setWeight(75)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.textEdit = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit.setGeometry(QtCore.QRect(30, 130, 201, 41))
        self.textEdit.setObjectName("textEdit")
        self.textEdit_2 = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_2.setGeometry(QtCore.QRect(30, 180, 201, 41))
        self.textEdit_2.setObjectName("textEdit_2")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(240, 180, 111, 41))
        self.pushButton_2.clicked.connect(self.saveDir) 


        font = QtGui.QFont()
        font.setFamily("Agency FB")
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(30, 300, 321, 71))
        self.pushButton_3.clicked.connect(self.getFile)

        font = QtGui.QFont()
        font.setFamily("Consolas")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_3.setFont(font)
        self.pushButton_3.setObjectName("pushButton_3")
        self.textEdit_3 = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_3.setGeometry(QtCore.QRect(150, 230, 201, 41))
        self.textEdit_3.setObjectName("textEdit_3")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(30, 240, 121, 21))
        font = QtGui.QFont()
        font.setFamily("Agency FB")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 385, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
    

    def fileUpload(self):
        fname=QFileDialog.getOpenFileName()
        if fname==False:
            return
        self.textEdit.setText(fname[0])

    def saveDir(self):
        fname=QFileDialog.getExistingDirectory()
        if fname == False:
            return
        self.textEdit_2.setText(fname)

    def getFile(self):
        try:
            m=mgrs.MGRS()
            file_path=self.textEdit.toPlainText().strip()
            save_path=self.textEdit_2.toPlainText().strip()
            sheet_name=self.textEdit_3.toPlainText().strip()
            if(file_path =='' or save_path=='' or sheet_name==''):
                QMessageBox.about(self,"알림창","공백란을 반드시 입력하세요.") 
                return
            
            data=pd.read_excel(file_path,sheet_name=sheet_name)
            data=data.fillna('')
            for i in range(len(data['군 MGRS'])):
                if data['군 MGRS'][i]:
                    if not '52S' in data['군 MGRS'][i]:
                        data['군 MGRS'][i]='52S' + data['군 MGRS'][i]
                    data["군 MGRS"][i] = data['군 MGRS'][i].replace(" ","")
                    data['GPS'][i]=m.toLatLon(data['군 MGRS'][i].encode())
            
            data.to_excel(save_path+'/'+sheet_name+'.xlsx')
            
            QMessageBox.about(self,"알림창","변환파일이 지정된 경로에 저장되었습니다.") 
            self.textEdit.setPlainText('')
            self.textEdit_2.setPlainText('')
            self.textEdit_3.setPlainText('')
        except Exception as e:
            QMessageBox.about(self,"알림창","에러내용: "+str(e)+"\n"+"파일 헤더에 군 MGRS, GPS 두 항목이 반드시 있어야합니다.\n 시트명 대소문자 구분하셔야 합니다.")


    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label.setText(_translate("MainWindow", "MGRS CONVERTER"))
        self.label_2.setText(_translate("MainWindow", "Release by JunHyeong"))
        self.pushButton.setText(_translate("MainWindow", "파일업로드"))
        self.pushButton_2.setText(_translate("MainWindow", "저장 경로 선택"))
        self.pushButton_3.setText(_translate("MainWindow", "변환 파일 저장"))
        self.label_3.setText(_translate("MainWindow", "엑셀 시트 이름"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

