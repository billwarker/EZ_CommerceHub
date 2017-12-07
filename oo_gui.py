import sys
from PyQt5.QtWidgets import (QMainWindow, QPushButton, QLabel,
	QFrame, QFileDialog, QApplication, QAction)
from PyQt5.QtGui import (QIcon, QColor)
from PyQt5.QtCore import QObject, pyqtSignal
import os
import oo_process
#from PyQt5.QtWidgets import *
#from PyQt5.QtGui import *

class OrderOpener(QMainWindow):

	def __init__(self):
		super().__init__()

		self.sheetCheck = False
		self.grouponPath = None
		self.commercehubPath = None
		self.staplesPath = None
		self.initUI()

	def initUI(self):

		menubar = self.menuBar()
		optionMenu = menubar.addMenu('Options')

		statusbar = self.statusBar()
		statusbar.showMessage('No sheets added.')

		grey = QColor(200, 200, 200)

		self.grouponSquare = QFrame(self)
		self.grouponSquare.setStyleSheet("QWidget { background-color: %s}" %
			grey.name())
		self.grouponSquare.setGeometry(50, 65, 80, 80)
		self.grouponSquare.mouseReleaseEvent = self.loadGroupon
		
		grouponLabel = QLabel('Groupon', self)
		grouponLabel.move(60, 125)


		self.commercehubSquare = QFrame(self)
		self.commercehubSquare.setStyleSheet("QWidget { background-color: %s}" %
			grey.name())
		self.commercehubSquare.setGeometry(200, 65, 80, 80)
		self.commercehubSquare.mouseReleaseEvent = self.loadCommerceHub
		commercehubLabel = QLabel('CommerceHub', self)
		commercehubLabel.move(210, 125)
		
		self.staplesSquare = QFrame(self)
		self.staplesSquare.setStyleSheet("QWidget { background-color: %s}" %
			grey.name())
		self.staplesSquare.setGeometry(370, 65, 80, 80)
		self.staplesSquare.mouseReleaseEvent = self.loadStaples
		staplesLabel = QLabel('Staples', self)
		staplesLabel.move(400, 125)


		processButton = QPushButton("Process", self)
		processButton.move(265, 200)
		processButton.clicked.connect(self.processing)

		exitButton = QPushButton("Exit", self)
		exitButton.move(375, 200)
		exitButton.clicked.connect(self.close)


		self.setGeometry(300, 300, 500, 250)
		self.setWindowTitle('Order Opener')
		self.show()

	def loadGroupon(self, event):

		fname = QFileDialog.getOpenFileName(self, 'Open file', '/')

		if fname[0]:
			var_path = fname[0]
			self.sheetCheck = True
			self.grouponPath = os.path.abspath(var_path)
			self.statusBar().showMessage('Groupon added!')


	def loadCommerceHub(self, event):

		fname = QFileDialog.getOpenFileName(self, 'Open file', '/')

		if fname[0]:
			var_path = fname[0]
			self.sheetCheck = True
			self.commercehubPath = os.path.abspath(var_path)
			self.statusBar().showMessage('CommerceHub added!')

	def loadStaples(self, event):

		fname = QFileDialog.getOpenFileName(self, 'Open file', '/')

		if fname[0]:
			var_path = fname[0]
			self.sheetCheck = True
			self.staplesPath = os.path.abspath(var_path)
			self.statusBar().showMessage('Staples added!')

	def processing(self):
		if self.sheetCheck:
			self.statusBar().showMessage('Processing sheet...')
			oo_process.process_output(self.grouponPath, self.commercehubPath,
				self.staplesPath)
			self.statusBar().showMessage('Done!')

		else:
			self.statusBar().showMessage('No sheets to process!')

if __name__ == '__main__':
	
	app = QApplication(sys.argv)
	run = OrderOpener()
	sys.exit(app.exec_())

