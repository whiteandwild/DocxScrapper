# PyQt5 Drag-and-Drop GUI Demo for MS Word Documents (demo.py)

import docx , sys

import os
from PyQt5 import QtWidgets, QtCore
import konwerter

# pip install pyinstaller
# pip install python-docx
# pip install PyQt5
# pip install pypiwin32

config_name = 'konwerter.py'
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)

config_path = os.path.join(application_path, config_name)
try:
    os.mkdir(os.path.join(application_path , 'Output'))
except:
    print("Dir already exists")
os.chdir((os.path.join(application_path , 'Output')))


def function1(target , nazwa):
    il = konwerter.convert(target , nazwa)
    os.chdir((os.path.join(application_path , 'Output')))
    return [True , il]


class MainApplication(QtWidgets.QWidget):
    def __init__(self):
        super(MainApplication, self).__init__()
        self.setObjectName("MainApplication")
        self.resize(370, 307)
  
        self.pushButton = QtWidgets.QPushButton(self)
        self.pushButton.setGeometry(QtCore.QRect(20, 70, 331, 41))
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.load_file_but)
        self.textBrowser = QtWidgets.QTextBrowser(self)
        self.textBrowser.setGeometry(QtCore.QRect(20, 140, 331, 131))
        self.textBrowser.setObjectName("textBrowser")
        self.setWindowTitle(QtWidgets.QApplication.translate("MainApplication", "Main Application", None))
        self.pushButton.setText(QtWidgets.QApplication.translate("MainApplication", "Wybierz plik", None))
        self.textBrowser.setHtml(QtWidgets.QApplication.translate("MainApplication","<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.5pt; font-weight:400; font-style:normal;\">\n"

"<p align=\"center\" style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-size:8pt;\"><br /></p>\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Upusc plik tutaj</span></p></body></html>", None))

        # Enable dragging and dropping onto the GUI
        self.setAcceptDrops(True)
        self.show()

    def load_file_but(self):
        """
        Open a File dialog when the button is pressed
        :return:
        """

        # Get the file location
        self.fname = QtWidgets.QFileDialog.getOpenFileName(self, 'Open file')
        # Load the file from the location
        self.load_file()

    def load_file(self):
        targetfilename = self.fname[0].lower()
        target = targetfilename
        if not target : return
       
        if targetfilename[-5:len(targetfilename)]!='.docx': # If the file does not have a doc, docx, or txt extension, it will be rejected.
            self.errorInvalidfilename()
            return
        name = ""
        while True:

            name = str(self.takeinput())
            if name == 'False': return
            if name == '': QtWidgets.QMessageBox.information(self, "Warning", "niepoprawna nazwa")
            else: break
               

        doc = docx.Document(targetfilename)
        if target:
            out = function1(target , name)
            if out[0] : QtWidgets.QMessageBox.information(self,"Information", "zakonczono pomyslnie \nbłędów : " + str(out[1]))
    
    # The following three methods set up dragging and dropping for the app
    def takeinput(self):
        name , done = QtWidgets.QInputDialog.getText(
             self, 'Input Dialog', 'Podaj nazwe docelowego folderu:') 
        if done:
            return name
        else:
            return False

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls:
            e.accept()
        else:
            e.ignore()

    def dragMoveEvent(self, e):
        if e.mimeData().hasUrls:
            e.accept()
        else:
            e.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls:
            event.setDropAction(QtCore.Qt.CopyAction)
            event.accept()
            self.fname = []
            for url in event.mimeData().urls():
                self.fname.append(str(url.toLocalFile()))
            self.load_file()
        else:
            event.ignore()
            
    # The following are the functions for the warning boxes that come up upon triggering an error.
 
         
    def errorFileload(self):
         messagetext = 'File loading error.'
         QtWidgets.QMessageBox.information(self, "Warning", messagetext)
         return
         
    def errorInvalidfilename(self):
         messagetext = 'Invalid file.'
         QtWidgets.QMessageBox.information(self, "Warning", messagetext)
         return

    def errorpywin32(self):
         messagetext = 'Microsoft Word COM object not identified in registry.'
         QtWidgets.QMessageBox.information(self, "Information", messagetext)
         return

app = QtWidgets.QApplication(sys.argv)
form = MainApplication()
form.show()
app.exec_()

