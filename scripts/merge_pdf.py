# Jun Chen
# This is great tool.
# 10-2-2018
import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from merge import Ui_Dialog
import os, sys, os.path, io,glob
from PyPDF2 import PdfFileMerger,PdfFileReader, PdfFileWriter
#import comtypes.client
import comtypes
from pdf2image import convert_from_path, convert_from_bytes
import mammoth

class MyFirstGuiProgram(Ui_Dialog):
	def __init__(self, dialog):
		Ui_Dialog.__init__(self)
		self.setupUi(dialog)
		self.pushButton.clicked.connect(self.showDialog1)
		self.pushButton_2.clicked.connect(self.showDialog2)
		self.pushButton_3.clicked.connect(self.showDialog3)
		self.pushButton_4.clicked.connect(self.showDialog4)
		self.pushButton_10.clicked.connect(self.showDialog5)
		self.pushButton_6.clicked.connect(self.convert1)
		self.pushButton_5.clicked.connect(self.convert2)
		self.pushButton_7.clicked.connect(self.convert3)
		self.pushButton_8.clicked.connect(self.convert4)
		self.pushButton_9.clicked.connect(self.convert5)
		self.pushButton_11.clicked.connect(self.clickBox5)
	def convert1(self):
		foldernameXX = str(self.textEdit.toPlainText())
		os.chdir(foldernameXX)
		pdfs = glob.glob('*.pdf')
		merger = PdfFileMerger()
		for pdf in pdfs:
			merger.append(open(pdf, 'rb'))
			with open('merged.pdf', 'wb') as fout:
				merger.write(fout)
		QtWidgets.QMessageBox.about(None, "INFO", "Finished")
#	self.results_summary.setText("Finished, Please check !!!")
	def convert2(self):
		foldernameXX = str(self.textEdit_2.toPlainText())
		path=os.path.dirname(foldernameXX)
		os.chdir(path)
		inputpdf = PdfFileReader(open(foldernameXX, "rb"))
#		pages_print=[1,3,5]   # need to edit this one later.
		pages_print=str(self.textEdit_5.toPlainText()).split(",")   # need to edit this one later."42,0".split(",")
		for i in pages_print:
			output = PdfFileWriter()
			j=int(i)-1
			output.addPage(inputpdf.getPage(j))
			with open("document-page%s.pdf" % i, "wb") as outputStream:
				output.write(outputStream)
		QtWidgets.QMessageBox.about(None, "INFO", "Finished")  # split file which need this file was not encriptyed. 
	def convert3(self):
		foldernameXX = str(self.textEdit_3.toPlainText())
		path=os.path.dirname(foldernameXX)
		os.chdir(path)
		wdFormatPDF = 17
		out_file=(os.path.splitext(foldernameXX)[0]+".pdf")		
		word = comtypes.client.CreateObject('Word.Application')
		doc = word.Documents.Open(foldernameXX)
		doc.SaveAs(out_file, FileFormat=wdFormatPDF)
		doc.Close()
		word.Quit()
		QtWidgets.QMessageBox.about(None, "INFO", "Finished")
	def convert4(self):
		foldernameXX = str(self.textEdit_4.toPlainText())
		path=os.path.dirname(foldernameXX)
		os.chdir(path)
		out_file=(os.path.splitext(foldernameXX)[0]+".html")		
		f = open(foldernameXX, 'rb')
		b = open(out_file, 'wb')
		document = mammoth.convert_to_html(f)
		b.write(document.value.encode('utf8'))
		f.close()
		b.close()
		QtWidgets.QMessageBox.about(None, "INFO", "Finished")
		
	def convert5(self):
		foldernameXX = str(self.textEdit_6.toPlainText())
		path=os.path.dirname(foldernameXX)
		os.chdir(path)
		images = convert_from_path(foldernameXX,500)
#		images[0].save('test.png', 'png')
#		images[1].save('test1.png', 'png')
		pages_print=str(self.textEdit_7.toPlainText()).split(",")   # need to edit this one later."42,0".split(",")
		for i in pages_print:
			j=int(i)-1
			images[j].save("image%s.png" % i, 'png')
		QtWidgets.QMessageBox.about(None, "INFO", "Finished")		
	def showDialog1(self):
		fname = QtWidgets.QFileDialog.getExistingDirectory(None, "Select Directory")
		if fname:
			self.textEdit.setText(fname)
	def showDialog2(self):
		fname,_ = QtWidgets.QFileDialog.getOpenFileName(None, "Open File", '.', "(*.pdf)")
		if fname:
			self.textEdit_2.setText(fname)
	def showDialog3(self):
		fname,_ = QtWidgets.QFileDialog.getOpenFileName(None, "Open File", '.', "(*.doc *.docx)")
		if fname:
			self.textEdit_3.setText(fname)
	def showDialog4(self):
		fname,_ = QtWidgets.QFileDialog.getOpenFileName(None, "Open File", '.', "(*.doc *.docx)")
		if fname:
			self.textEdit_4.setText(fname)
	def showDialog5(self):
		fname,_ = QtWidgets.QFileDialog.getOpenFileName(None, "Open File", '.', "(*.pdf)")
		if fname:
			self.textEdit_6.setText(fname)
	def clickBox5(self):
		QtWidgets.QMessageBox.about(None, "About", "Author: <i>.Jun Chen</i>")
#https://pythonspot.com/pyqt5-file-dialog/			
if __name__ == '__main__':
	app = QtWidgets.QApplication(sys.argv)
	dialog = QtWidgets.QDialog()
	prog = MyFirstGuiProgram(dialog)
	dialog.show()
	sys.exit(app.exec_())
