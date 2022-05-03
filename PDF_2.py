from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import *
import os, PyPDF2, webbrowser
from docx2pdf import convert


class Ui_MainWindow(object):

    def __init__(self):
        self.selected = []
        self.selected_pdf_path = []
        self.count = 0

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 591)
        MainWindow.setStyleSheet("border: 0px;")

        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.top_buttons_frame = QtWidgets.QFrame(self.centralwidget)
        self.top_buttons_frame.setGeometry(QtCore.QRect(0, 0, 801, 101))
        self.top_buttons_frame.setStyleSheet("background: #ad2f2f;")
        self.top_buttons_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.top_buttons_frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.top_buttons_frame.setObjectName("top_buttons_frame")

        self.remove_button = QtWidgets.QToolButton(self.top_buttons_frame)
        self.remove_button.setGeometry(QtCore.QRect(150, 10, 71, 71))
        self.remove_button.setStyleSheet("background: #ad2f2f; border: 0px;")
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/images/icons8-delete-128.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.remove_button.setIcon(icon)
        self.remove_button.setIconSize(QtCore.QSize(100, 100))
        self.remove_button.setObjectName("remove_button")
        self.remove_button.clicked.connect(self.remove)
        
        self.remove_label = QtWidgets.QLabel(self.top_buttons_frame)
        self.remove_label.setGeometry(QtCore.QRect(150, 70, 71, 21))
        self.remove_label.setStyleSheet("font-size: 20px; font-family: forte;")
        self.remove_label.setAlignment(QtCore.Qt.AlignCenter)
        self.remove_label.setObjectName("remove_label")
        
        self.up_button = QtWidgets.QToolButton(self.top_buttons_frame)
        self.up_button.setGeometry(QtCore.QRect(260, 10, 71, 71))
        self.up_button.setStyleSheet("background: #ad2f2f; border: 0px;")
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(":/images/icons8-send-letter-96 (2).png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.up_button.setIcon(icon1)
        self.up_button.setIconSize(QtCore.QSize(100, 100))
        self.up_button.setObjectName("up_button")
        self.up_button.clicked.connect(self.up)
        
        self.up_label = QtWidgets.QLabel(self.top_buttons_frame)
        self.up_label.setGeometry(QtCore.QRect(270, 70, 51, 21))
        self.up_label.setStyleSheet("font-size: 20px; font-family: forte;")
        self.up_label.setAlignment(QtCore.Qt.AlignCenter)
        self.up_label.setObjectName("up_label")
        
        self.down_button = QtWidgets.QToolButton(self.top_buttons_frame)
        self.down_button.setGeometry(QtCore.QRect(370, 10, 71, 71))
        self.down_button.setStyleSheet("background: #ad2f2f; border: 0px;")
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap(":/images/icons8-low-importance-96 (1).png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.down_button.setIcon(icon2)
        self.down_button.setIconSize(QtCore.QSize(100, 100))
        self.down_button.setObjectName("down_button")
        self.down_button.clicked.connect(self.down)
        
        self.down_label = QtWidgets.QLabel(self.top_buttons_frame)
        self.down_label.setGeometry(QtCore.QRect(380, 70, 51, 21))
        self.down_label.setStyleSheet("font-size: 20px; font-family: forte;")
        self.down_label.setAlignment(QtCore.Qt.AlignCenter)
        self.down_label.setObjectName("down_label")
        
        self.add_button = QtWidgets.QToolButton(self.top_buttons_frame)
        self.add_button.setGeometry(QtCore.QRect(40, 10, 71, 71))
        self.add_button.setStyleSheet("background: #ad2f2f; border: 0px;")
        icon3 = QtGui.QIcon()
        icon3.addPixmap(QtGui.QPixmap(":/images/icons8-plus-128 (1).png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.add_button.setIcon(icon3)
        self.add_button.setIconSize(QtCore.QSize(60, 60))
        self.add_button.setObjectName("add_button")
        self.add_button.clicked.connect(self.select_pdf)
        
        self.add_label = QtWidgets.QLabel(self.top_buttons_frame)
        self.add_label.setGeometry(QtCore.QRect(30, 70, 91, 21))
        self.add_label.setStyleSheet("font-size: 20px; font-family: forte;")
        self.add_label.setAlignment(QtCore.Qt.AlignCenter)
        self.add_label.setObjectName("add_label")
        
        self.heading_label = QtWidgets.QLabel(self.top_buttons_frame)
        self.heading_label.setGeometry(QtCore.QRect(530, 20, 261, 71))
        self.heading_label.setStyleSheet("font-family: Cooper Black; font-size: 30px; color: #d8e8e6;")
        self.heading_label.setAlignment(QtCore.Qt.AlignCenter)
        self.heading_label.setWordWrap(True)
        self.heading_label.setObjectName("heading_label")
        
        self.icon_frame = QtWidgets.QFrame(self.centralwidget)
        self.icon_frame.setGeometry(QtCore.QRect(540, 100, 261, 471))
        self.icon_frame.setStyleSheet("background: #ad2f2f;")
        self.icon_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.icon_frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.icon_frame.setObjectName("icon_frame")
        
        self.merge_button = QtWidgets.QToolButton(self.icon_frame)
        self.merge_button.setGeometry(QtCore.QRect(40, 170, 71, 71))
        self.merge_button.setStyleSheet("background: #ad2f2f; border: 0px;")
        icon4 = QtGui.QIcon()
        icon4.addPixmap(QtGui.QPixmap(":/images/icons8-merge-files-80.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.merge_button.setIcon(icon4)
        self.merge_button.setIconSize(QtCore.QSize(60, 60))
        self.merge_button.setObjectName("merge_button")
        self.merge_button.clicked.connect(self.merge)
        
        self.merge_label = QtWidgets.QLabel(self.icon_frame)
        self.merge_label.setGeometry(QtCore.QRect(40, 240, 71, 31))
        self.merge_label.setStyleSheet("font-size: 20px; font-family: forte; color: yellow;")
        self.merge_label.setAlignment(QtCore.Qt.AlignCenter)
        self.merge_label.setObjectName("merge_label")
        
        self.convert_button = QtWidgets.QToolButton(self.icon_frame)
        self.convert_button.setGeometry(QtCore.QRect(160, 180, 51, 61))
        self.convert_button.setStyleSheet("background: #ad2f2f; border: 0px;")
        icon5 = QtGui.QIcon()
        icon5.addPixmap(QtGui.QPixmap(":/images/convert_png_to_icon_320554_44544.jpg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.convert_button.setIcon(icon5)
        self.convert_button.setIconSize(QtCore.QSize(60, 60))
        self.convert_button.setObjectName("convert_button")
        self.convert_button.clicked.connect(self.docx_to_pdf)
        
        self.convert_label = QtWidgets.QLabel(self.icon_frame)
        self.convert_label.setGeometry(QtCore.QRect(150, 240, 71, 31))
        self.convert_label.setStyleSheet("font-size: 20px; font-family: forte; color: yellow;")
        self.convert_label.setAlignment(QtCore.Qt.AlignCenter)
        self.convert_label.setObjectName("convert_label")
        
        self.icon = QtWidgets.QToolButton(self.icon_frame)
        self.icon.setGeometry(QtCore.QRect(60, 20, 151, 121))
        self.icon.setStyleSheet("border: 0px")
        icon6 = QtGui.QIcon()
        icon6.addPixmap(QtGui.QPixmap(":/images/icon.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.icon.setIcon(icon6)
        self.icon.setIconSize(QtCore.QSize(300, 300))
        self.icon.setObjectName("icon")
        self.icon.clicked.connect(self.developer)
        
        self.clear_button = QtWidgets.QToolButton(self.icon_frame)
        self.clear_button.setGeometry(QtCore.QRect(100, 290, 61, 51))
        icon7 = QtGui.QIcon()
        icon7.addPixmap(QtGui.QPixmap(":/images/icons8-broom-96.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.clear_button.setIcon(icon7)
        self.clear_button.setIconSize(QtCore.QSize(60, 60))
        self.clear_button.setObjectName("clear_button")
        self.clear_button.clicked.connect(self.clear)
        
        self.merge_label_2 = QtWidgets.QLabel(self.icon_frame)
        self.merge_label_2.setGeometry(QtCore.QRect(90, 340, 71, 31))
        self.merge_label_2.setStyleSheet("font-size: 20px; font-family: forte; color: yellow;")
        self.merge_label_2.setAlignment(QtCore.Qt.AlignCenter)
        self.merge_label_2.setObjectName("merge_label_2")
        
        self.insta = QtWidgets.QToolButton(self.icon_frame)
        self.insta.setGeometry(QtCore.QRect(50, 410, 31, 31))
        icon8 = QtGui.QIcon()
        icon8.addPixmap(QtGui.QPixmap(":/images/icons8-instagram-48.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.insta.setIcon(icon8)
        self.insta.setIconSize(QtCore.QSize(50, 50))
        self.insta.setObjectName("insta")
        self.insta.clicked.connect(self.openinsta)

        self.fb = QtWidgets.QToolButton(self.icon_frame)
        self.fb.setGeometry(QtCore.QRect(110, 410, 31, 31))
        icon9 = QtGui.QIcon()
        icon9.addPixmap(QtGui.QPixmap(":/images/icons8-facebook-48.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.fb.setIcon(icon9)
        self.fb.setIconSize(QtCore.QSize(50, 50))
        self.fb.setObjectName("fb")
        self.fb.clicked.connect(self.openfb)
        
        self.gmail = QtWidgets.QToolButton(self.icon_frame)
        self.gmail.setGeometry(QtCore.QRect(180, 410, 31, 31))
        self.gmail.setStyleSheet("background: white;")
        icon10 = QtGui.QIcon()
        icon10.addPixmap(QtGui.QPixmap(":/images/icons8-gmail-48.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.gmail.setIcon(icon10)
        self.gmail.setIconSize(QtCore.QSize(50, 50))
        self.gmail.setObjectName("gmail")
        self.gmail.clicked.connect(self.opengmail)
        
        self.display_frame = QtWidgets.QFrame(self.centralwidget)
        self.display_frame.setGeometry(QtCore.QRect(0, 100, 541, 471))
        self.display_frame.setStyleSheet("background: white;")
        self.display_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.display_frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.display_frame.setObjectName("display_frame")
        
        self.tableWidget = QtWidgets.QTableWidget(self.display_frame)
        self.tableWidget.setGeometry(QtCore.QRect(0, 10, 541, 410))
        self.tableWidget.setStyleSheet("border: 2px; font-size: 15px; font-family: times new roman; color: blue;")
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(2)
        self.tableWidget.setRowCount(1)
        self.tableWidget.horizontalHeader().setDefaultSectionSize(260)
        self.tableWidget.clicked.connect(self.get_clicked_row_col)
        # For the horizontal labels to be visible
        self.tableWidget.setHorizontalHeaderLabels(["File Name", "Path"])
        self.tableWidget.horizontalHeader().setStyleSheet("font-size: 16px; font-family: elephant; color: black;")
        self.tableWidget.setRowCount(0)

        self.progressbar = QProgressBar(self.display_frame)
        self.progressbar.setGeometry(QtCore.QRect(30, 430, 470, 25))
        self.progressbar.setStyleSheet("border: 1px solid black; font-size: 15px; font-family: times new roman; text-align: center;")
        self.progressbar.setValue(0)
        
        MainWindow.setCentralWidget(self.centralwidget)
        
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.remove_button.setText(_translate("MainWindow", "..."))
        self.remove_label.setText(_translate("MainWindow", "Remove"))
        self.up_button.setText(_translate("MainWindow", "..."))
        self.up_label.setText(_translate("MainWindow", "Up"))
        self.down_button.setText(_translate("MainWindow", "..."))
        self.down_label.setText(_translate("MainWindow", "Down"))
        self.add_button.setText(_translate("MainWindow", "..."))
        self.add_label.setText(_translate("MainWindow", "Add Files"))
        self.heading_label.setText(_translate("MainWindow", "PDF Merger and Converter"))
        self.merge_button.setText(_translate("MainWindow", "..."))
        self.merge_label.setText(_translate("MainWindow", "Merge"))
        self.convert_button.setText(_translate("MainWindow", "..."))
        self.convert_label.setText(_translate("MainWindow", "Convert"))
        self.icon.setText(_translate("MainWindow", "..."))
        self.clear_button.setText(_translate("MainWindow", "..."))
        self.merge_label_2.setText(_translate("MainWindow", "Clear"))
        self.insta.setText(_translate("MainWindow", "..."))
        self.fb.setText(_translate("MainWindow", "..."))
        self.gmail.setText(_translate("MainWindow", "..."))

    def select_pdf(self):
        file = QFileDialog.getOpenFileNames(filter = "*.pdf *.docx")
        if file[0]:
            if not self.selected:
                for i in file[0]:
                    self.selected.append(os.path.basename(i))
                    self.selected_pdf_path.append(i)
                self.if_selected_pdf_empty()
            else:
                for i in file[0]:
                    self.selected.append(os.path.basename(i))
                    self.selected_pdf_path.append(i)
                self.if_selected_pdf_nonempty(file[0])


    def if_selected_pdf_empty(self):
        for j in range(len(self.selected)):
            if self.tableWidget.columnCount() == 0:
                self.tableWidget.insertRow(self.count)
                self.tableWidget.insertColumn(0)
                self.tableWidget.insertColumn(1)
            else:
                self.tableWidget.insertRow(self.count)

            self.tableWidget.setItem(self.count, 0, QTableWidgetItem(self.selected[self.count]))
            self.tableWidget.setItem(self.count, 1, QTableWidgetItem(self.selected_pdf_path[self.count]))
            self.count += 1

    def if_selected_pdf_nonempty(self, file):
        for j in range(len(file)):
            self.tableWidget.insertRow(self.count)
            self.tableWidget.setItem(self.count, 0, QTableWidgetItem(self.selected[self.count]))
            self.tableWidget.setItem(self.count, 1, QTableWidgetItem(self.selected_pdf_path[self.count]))
            self.count += 1

    def get_clicked_row_col(self, item):
        self.clicked_row = item.row()

    def up(self):
        try:
            if self.clicked_row > 0:
                selected_temp = [self.tableWidget.item(self.clicked_row, 0).text(), self.tableWidget.item(self.clicked_row, 1).text()]
                above_temp = [self.tableWidget.item(self.clicked_row - 1, 0).text(), self.tableWidget.item(self.clicked_row - 1, 1).text()]

                self.tableWidget.setItem(self.clicked_row, 0, QTableWidgetItem(above_temp[0]))
                self.tableWidget.setItem(self.clicked_row, 1, QTableWidgetItem(above_temp[1]))
                self.tableWidget.setItem(self.clicked_row - 1, 0, QTableWidgetItem(selected_temp[0]))
                self.tableWidget.setItem(self.clicked_row - 1, 1, QTableWidgetItem(selected_temp[1]))

                self.selected[self.clicked_row], self.selected[self.clicked_row - 1] = self.selected[self.clicked_row - 1], self.selected[self.clicked_row]
                self.selected_pdf_path[self.clicked_row], self.selected_pdf_path[self.clicked_row - 1] = self.selected_pdf_path[self.clicked_row - 1], self.selected_pdf_path[self.clicked_row]
        except:
            pass

    def down(self):
        try:
            if self.clicked_row < len(self.selected) - 1:
                selected_temp = [self.tableWidget.item(self.clicked_row, 0).text(), self.tableWidget.item(self.clicked_row, 1).text()]
                below_temp = [self.tableWidget.item(self.clicked_row + 1, 0).text(), self.tableWidget.item(self.clicked_row + 1, 1).text()]

                self.tableWidget.setItem(self.clicked_row, 0, QTableWidgetItem(below_temp[0]))
                self.tableWidget.setItem(self.clicked_row, 1, QTableWidgetItem(below_temp[1]))
                self.tableWidget.setItem(self.clicked_row + 1, 0, QTableWidgetItem(selected_temp[0]))
                self.tableWidget.setItem(self.clicked_row + 1, 1, QTableWidgetItem(selected_temp[1]))

                self.selected[self.clicked_row], self.selected[self.clicked_row + 1] = self.selected[self.clicked_row + 1], self.selected[self.clicked_row]
                self.selected_pdf_path[self.clicked_row], self.selected_pdf_path[self.clicked_row + 1] = self.selected_pdf_path[self.clicked_row + 1], self.selected_pdf_path[self.clicked_row]
        except:
            pass

    def remove(self):
        if self.selected:
            self.tableWidget.removeRow(self.clicked_row)
            del self.selected[self.clicked_row]
            del self.selected_pdf_path[self.clicked_row]
            self.count -= 1
            if not self.selected:
                self.count = 0

    def clear(self):
        self.tableWidget.setRowCount(0)
        self.selected = []
        self.selected_pdf_path = []
        self.count = 0

    def merge(self):
        if self.selected_pdf_path:
                
            output_file = QFileDialog.getSaveFileName()[0]
            
            if output_file:
                self.progressbar.setValue(1)
                output_file += ".pdf"
                i = 1

                pdfmerger = PyPDF2.PdfFileMerger()

                #loop through all PDFs
                for file in self.selected_pdf_path:
                    if file.endswith("docx"):
                        file = file.replace(".docx", ".pdf")
                        convert(file[:-3] + "docx", file)
                    pdfmerger.append(PyPDF2.PdfFileReader(open(file, "rb")))
                    self.value = ((100 * i) // (len(self.selected) + 1))
                    self.progressbar.setValue(self.value)
                    i += 1

                
                pdfmerger.write(output_file)
                self.progressbar.setValue(100)
                i = 0
                self.messagebox("SUCCESS", "PDFs were successfully merged. Please visit the destination directory.", "ic")
        
        else:
            self.messagebox("FILE NOT SELECTED", "Please select atleast one file.", "w")

    def close_progress(self):
        self.progressbar.setValue(0)


    def openfb(self):
        webbrowser.open_new(r"https://www.facebook.com/akashkumarsingh17272888/")

    def openinsta(self):
        webbrowser.open_new(r"https://www.instagram.com/pythonfriendly/")

    def opengmail(self):
        self.messagebox("e-mail ME AT", "Email at: akashkumar8462@gmail.com", "i")

    def developer(self):
        msg = "Developer: Akash Kumar Singh\nDate created: Feb 13, 2021"
        self.messagebox("Developer Details", msg, "i")

    def messagebox(self, title, text, icon):
        msgBox = QMessageBox()
        if icon == "w":
            msgBox.setIcon(QMessageBox.Warning)    
        else:
            msgBox.setIcon(QMessageBox.Information)
        msgBox.setText(text)
        msgBox.setWindowTitle(title)
        msgBox.setStandardButtons(QMessageBox.Ok)
        if icon == "ic":
            msgBox.buttonClicked.connect(self.close_progress)
        msgBox.exec()

    def docx_to_pdf(self):
        if self.selected:
            directory = QFileDialog.getExistingDirectory()
            if directory:
                self.progressbar.setValue(1)
                j = 1
                for i in self.selected_pdf_path:
                    name = os.path.basename(i)
                    if name.endswith(".docx"):
                        name = directory + "/" + name[:-4] + "pdf"
                        convert(i, name)
                    self.value = ((100 * j) // len(self.selected))
                    self.progressbar.setValue(self.value)
                    j += 1

                j = 0    
                self.messagebox("CONVERSION SUCCESSFUL", "Conversion successfully completed.\nPlease visit the destination folder.", "ic")
        else:
            self.messagebox("FILE NOT SELECTED", "Please select atleast one file.", "w")
    
import image
 

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_()) 
