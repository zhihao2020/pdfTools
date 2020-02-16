from ui import Ui_Form
from PyQt5.QtWidgets import *
import sys
import PyPDF2
import os
import string
import itertools

class reload(QWidget,Ui_Form):
    def __init__(self):
        super(reload,self).__init__()
        self.setupUi(self)
        self.init()

    def init(self):
        # 变量
        self.filenames = []
        self.index = 0
        # 添加文件
        self.pushButton.clicked.connect(self.chaifen)
        self.pushButton_2.clicked.connect(self.hebing)
        self.pushButton_4.clicked.connect(self.pojie)
        self.pushButton_9.clicked.connect(self.jiami)
        # 绑定开始
        self.pushButton_5.clicked.connect(self.startPojie)
        self.pushButton_6.clicked.connect(self.startChaifen)
        self.pushButton_7.clicked.connect(self.startChaifen2)
        self.pushButton_3.clicked.connect(self.startHebing)
        self.pushButton_8.clicked.connect(self.startJiami)

    def chaifen(self):
        try:
            self.listWidget.clear()
            self.filenames= []
            filedialog = QFileDialog()
            filedialog.setFileMode(QFileDialog.ExistingFile)
            filedialog.setNameFilter("文本文件(*.pdf)")
            if filedialog.exec_():
                self.filenames = filedialog.selectedFiles()
                print(self.filenames)
            self.listWidget.addItem(str(self.filenames[0]))
            self.listWidget.show()
        except:pass

    def startChaifen(self):
        try:
            pdfFile = open(self.filenames[0], 'rb')

            pdfReader = PyPDF2.PdfFileReader(pdfFile)
            pdfWriter = PyPDF2.PdfFileWriter()
            start = self.spinBox.value()
            fin = self.spinBox_2.value()

            if int(start) > pdfReader.numPages or int(fin)>pdfReader.numPages:
                print(QMessageBox.information(self, '提示', '拆分页码大于总页数', QMessageBox.Yes, QMessageBox.Yes))

            else:
                if int(start) < int(fin):
                    print("进入拆分-----------")
                    for pageNum in range(int(start) , int(fin)):
                        pageObj = pdfReader.getPage(pageNum)
                        pdfWriter.addPage(pageObj)

                    pdfOutputFile = open(os.getcwd()+"\\拆分.pdf", "wb")
                    pdfWriter.write(pdfOutputFile)

                    pdfOutputFile.close()
                    pdfFile.close()
                    print(QMessageBox.information(self, '提示', '拆分成功\n文件在程序目录下', QMessageBox.Yes,QMessageBox.Yes))
                else:
                    print(QMessageBox.information(self, '提示', '结束页码 小于 开始页码', QMessageBox.Yes, QMessageBox.Yes))

        except IndexError:

            print((QMessageBox.information(self, '提示', '请先添加PDF', QMessageBox.Yes,
                                      QMessageBox.Yes)))
    def startChaifen2(self):
        try:
            pdfFile = open(self.filenames[0], 'rb')
        except IndexError:
            print((QMessageBox.information(self, '提示', '请先添加PDF', QMessageBox.Yes,
                                      QMessageBox.Yes)))
        pdfReader = PyPDF2.PdfFileReader(pdfFile)
       # pdfWriter = PyPDF2.PdfFileWriter()
        fenshu = self.spinBox_3.value()
        print("选择：", fenshu)
        zhangshu = pdfReader.numPages // fenshu
        pan = 1
        for n in range(fenshu):
            if pan == fenshu:
                print("进入拆分最后结尾...")
                start = zhangshu * (pan-1)
                fin =pdfReader.numPages
                print("开始页码：", start)
                print("结束页码：", fin)
                pdfriter = self.startChaifen2_1(pdfReader, start, fin)
                pdfOutFile = open(str(os.getcwd()) +"\\"+str(pan) + '.pdf', "wb")
                print("输出：", str(os.getcwd()) + "\\" + str(pan) + '.pdf')
                pdfriter.write(pdfOutFile)
                pdfOutFile.close()
                pdfFile.close()

            else:
                start = zhangshu * (pan-1)
                fin = zhangshu * (pan)
                print("开始页码：", start)
                print("结束页码：", fin)
                pdfriter = self.startChaifen2_1(pdfReader,start,fin)
                pan += 1
                pdfOutFile = open(str(os.getcwd()) + "\\"+str(pan-1) + '.pdf', "wb")
                print("输出：",str(os.getcwd()) + "\\"+str(pan-1) + '.pdf')
                pdfriter.write(pdfOutFile)
                pdfOutFile.close()
        print((QMessageBox.information(self, '提示', '拆分成功\n文件在程序目录下', QMessageBox.Yes,
                                               QMessageBox.Yes)))
        pdfFile.close()

    def startChaifen2_1(self,name,start,fin):
        pdfWriter = PyPDF2.PdfFileWriter()
        for n in range(start,fin):
            pageObj = name.getPage(n)
            pdfWriter.addPage(pageObj)
        return pdfWriter

    def hebing(self):
        try:
            self.listWidget_2.clear()
            self.filenames = []
            filedialog = QFileDialog()
            filedialog.setFileMode(QFileDialog.ExistingFiles)
            filedialog.setNameFilter("文本文件(*.pdf)")

            if filedialog.exec_():
                self.filenames = filedialog.selectedFiles()
                for name in self.filenames:
                    self.listWidget_2.addItem(str(name))
            self.listWidget_2.show()
        except:pass
    def startHebing(self):
        pdfWriter = PyPDF2.PdfFileWriter()
        for name in self.filenames:
            print(name)
            pdfFile = open(name,"rb")
            pdfReader = PyPDF2.PdfFileReader(pdfFile, strict=False)
            for pageNum in range(pdfReader.numPages):
                pageObj = pdfReader.getPage(pageNum)
                pdfWriter.addPage(pageObj)
        ns = 1
        if f'合并{ns}.pdf' in os.listdir(os.getcwd()):
            #print(f'合并{ns}.pdf',"----",os.listdir(os.getcwd()))
            ns +=1
            pdfOutputFile = open(os.getcwd()+f'\\合并{ns}.pdf',"wb")
            pdfWriter.write(pdfOutputFile)
            pdfOutputFile.close()
            pdfFile.close()
        else:
            #print(f'合并{ns}.pdf', "----", os.listdir(os.getcwd()))
            pdfOutputFile = open(os.getcwd() + '\\合并1.pdf', "wb")
            pdfWriter.write(pdfOutputFile)
            pdfOutputFile.close()
            pdfFile.close()
        print(QMessageBox.information(self, '提示', '合并成功', QMessageBox.Yes,
                                  QMessageBox.Yes))

    def pojie(self):
        self.listWidget_3.clear()
        self.filenames = []
        filedialog = QFileDialog()
        filedialog.setFileMode(QFileDialog.ExistingFiles)
        filedialog.setNameFilter("文本文件(*.pdf)")

        if filedialog.exec_():
            self.filenames = filedialog.selectedFiles()
            print(self.filenames)

        self.listWidget_3.addItem(str(self.filenames))
        self.listWidget_3.show()


    def get_strings(self):
        char = string.printable
        strings = []
        for i in range(int(self.lineEdit_3.text()),int(self.lineEdit.text())+1):
            strings.append((itertools.product(char, repeat=i),))
        print(strings)
        return itertools.chain(*strings)


    def startPojie(self):
        pan = 0
        try:
            pdfFile = open(self.filenames[0],"rb")
            pdfReader = PyPDF2.PdfFileReader(pdfFile)
            pdfWriter = PyPDF2.PdfFileWriter()
            if pdfReader.isEncrypted:
                print("解密中ing......")
                list_str = self.get_strings()
                for x in list_str:
                    QApplication.processEvents()
                    for y in x:
                        QApplication.processEvents()
                        print("密码：","".join(y))
                        if pdfReader.decrypt("".join(y)) == 1:
                            print(QMessageBox.information(self, '提示', '解密成功，密码为%s'%"".join(y), QMessageBox.Yes,
                                                   QMessageBox.Yes))
                            pan =1
                            raise Exception('This is the error message.')

            else:print(QMessageBox.information(self, '提示', '此文件并未加密！', QMessageBox.Yes,
                                                   QMessageBox.Yes))
            #if pan == 0:
             #   print(QMessageBox.information(self, '提示', '解密失败,位数不够', QMessageBox.Yes,
              #                                     QMessageBox.Yes))
            if pan ==1:
                for pageNum in range(pdfReader.numPages):
                    pageObj = pdfReader.getPage(pageNum)
                    pdfWriter.addPage(pageObj)
                pdfOutFile = open(os.getcwd()+"\\"+"解密成果.pdf","wb")
                pdfWriter.write(pdfOutFile)
                pdfOutFile.close()
                pdfFile.close()

        except IndexError:
            print(QMessageBox.information(self, '提示', '请添加pdf文件', QMessageBox.Yes,
                                          QMessageBox.Yes))

        except:pass

    def jiami(self):
        try:
            self.listWidget_4.clear()
            self.filenames = []
            filedialog = QFileDialog()
            filedialog.setFileMode(QFileDialog.ExistingFile)
            filedialog.setNameFilter("文本文件(*.pdf)")
            if filedialog.exec_():
                self.filenames = filedialog.selectedFiles()
                print(self.filenames)
            self.listWidget_4.addItem(str(self.filenames[0]))
            self.listWidget_4.show()
        except:pass
    def startJiami(self):
        try:
            pdfFile = open(self.filenames[0], 'rb')
        except IndexError:
            print((QMessageBox.information(self, '提示', '请先添加PDF', QMessageBox.Yes,
                                      QMessageBox.Yes)))
        pdfReader = PyPDF2.PdfFileReader(pdfFile)
        pdfWriter = PyPDF2.PdfFileWriter()
        if pdfReader.isEncrypted:
            print(print((QMessageBox.information(self, '提示', '此文件已经加密', QMessageBox.Yes,
                                      QMessageBox.Yes))))
        else:
            for pageNum in range(pdfReader.numPages):
                pdfWriter.addPage(pdfReader.getPage(pageNum))
            pdfWriter.encrypt(self.lineEdit_2.text())

            resultPdf = open(os.getcwd()+"\\加密.pdf","wb")
            pdfWriter.write(resultPdf)
            resultPdf.close()
            pdfFile.close()
            print(QMessageBox.information(self, '提示', "加密成功", QMessageBox.Yes,
                                          QMessageBox.Yes))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWin = reload()
    myWin.show()
    sys.exit(app.exec_())