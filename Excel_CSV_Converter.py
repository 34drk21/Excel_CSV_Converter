import xlrd
import csv
import sys
import os
import math
import mojimoji
from PySide2.QtWidgets import *
from PySide2.QtGui import *
from PySide2.QtCore import *

class ListView(QListView):
	def __init__(self, *args, **kwargs):
		super(ListView, self).__init__(*args, **kwargs)

		self.setDragEnabled(False)
		self.setAcceptDrops(True)
		self.setDropIndicatorShown(True)

	def dragEnterEvent(self, event):
		if event.mimeData().hasUrls():
			event.accept()

		else:
			super(ListView, self).dragEnterEvent(event)

	def dragMoveEvent(self, event):
		if event.mimeData().hasUrls():
			event.accept()
		else:
			super(ListView, self).dragMoveEvent(event)

	def dropEvent(self, event):
		if event.mimeData().hasUrls():
			model = self.model()
			urls = event.mimeData().urls()
			for url in urls:
				filename = url.toLocalFile()
				print(type(url))
				item = QStandardItem(filename)
				model.appendRow(item)

			event.accept()
		else:
			super(ListView, self).dropEvent(event)

	def removeSelectedItem(self):
		model = self.model()
		selModel = self.selectionModel()

		while True:
			indexes = selModel.selectedIndexes()
			if not indexes:
				break
	             
			model.removeRow(indexes[0].row())
	             
	def keyPressEvent(self, event):
		if event.key() == Qt.Key_Delete or event.key() == Qt.Key_X:
			self.removeSelectedItem()
			return
		super(ListView, self).keyPressEvent(event)
class MainWindow(QWidget):
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		self.setupUI()


	def setupUI(self):

		self.title = "Excel to CSV Converter"

		layout = QVBoxLayout()

		filesLabel = QLabel()
		filesLabel.setText("Files To Convert: ")

		self.convertButton = QPushButton("Convert")
		self.convertButton.setFocusPolicy(Qt.StrongFocus)
		self.convertButton.clicked.connect(self.outputFile)

		self.model = QStandardItemModel()
		self.listWindow = ListView()
		self.listWindow.setWindowFlags(Qt.Window)
		self.listWindow.setModel(self.model)

		layout.addWidget(filesLabel)
		layout.addWidget(self.listWindow)
		layout.addWidget(self.convertButton)

		self.setLayout(layout)


	def loadFile(self, path):
		self.workbook = xlrd.open_workbook(path)
		self.sheet = self.workbook.sheet_by_index(0)

	def outputFile(self):
		#clean file name
		#urls = self.model.stringToList()
		model = self.listWindow.model()

		for i in range(model.rowCount()):

			self.loadFile(model.item(i).text())
			self.findStartRow()

			for sheet in self.workbook.sheets():
				sheet_name = sheet.name
				url = model.item(i).text().replace(".xls", ".csv")
				output =  url
				print(output)
				#output = "/Users/shotakimura/Documents/Python/excel/sample.csv"
				with open(output, 'w') as fp:
					writer = csv.writer(fp)
					writer.writerow(self.initCategory())

					i = self.startRow
					for i in range(self.startRow, self.sheet.nrows):
						writer.writerow(self.getItems(i))
						i = i + 1


			print("saved at:" + output)

	def cleanFilename(self, filename):
		fn = filename.split("/")
		fn = fn[-1]
		fn = fn.replace(".xls", "")
		return fn

	def initCategory(self):
		return ["レイアウト", "JANコード", "メーカー・産地", "商品名", "規格１", "売価１", "予備１"]

	def findStartRow(self):
		for i in range(self.sheet.nrows):
			if type(self.sheet.cell(i, 0).value) is float:
				self.startRow = i
				break

	def getItems(self, row):
		jan_code = str(self.sheet.row(row)[2].value).replace(".0", "")
		maker = self.sheet.row(row)[3].value
		name = self.sheet.row(row)[4].value
		standard = self.sheet.row(row)[5].value
		price = str(self.sheet.row(row)[10].value).replace(".0", "")
		try:
			comment = str(self.sheet.row(row)[21].value)
		except:
			comment = ""
		return ["", jan_code,
					mojimoji.han_to_zen(maker.rstrip()),
					mojimoji.han_to_zen(name.rstrip()),
					mojimoji.han_to_zen(str(standard).rstrip()),
					price,
					mojimoji.han_to_zen(comment.rstrip())
					]
		#items = ["", 55, 67]
		#return items


def main():
	app = QApplication(sys.argv)
	window = MainWindow()
	window.show()
	sys.exit(app.exec_())

if __name__=='__main__':
	main()