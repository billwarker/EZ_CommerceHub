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
		self.dark_grey = QColor(170, 170, 170)
		program_w = 500
		program_h = program_w / 2
		box_dim = 80
		label_height = 90
		box_h = 65
		box_offset = program_w / 10

		# Groupon
		self.grouponBox = QFrame(self)
		self.grouponBox.setStyleSheet("QWidget { background-color: %s}" %
			grey.name())
		self.grouponBox.setGeometry(box_offset, box_h, box_dim, box_dim)
		self.grouponLabel = QLabel('Groupon', self)
		self.grouponLabel.move(box_offset + 17, label_height)

		self.grouponBox.mouseReleaseEvent = self.loadGroupon
		self.grouponLabel.mouseReleaseEvent = self.loadGroupon

		# CommerceHub
		self.commerceBox = QFrame(self)
		self.commerceBox.setStyleSheet("QWidget { background-color: %s}" %
			grey.name())
		commerceBox_w = (program_w/2) - (box_dim/2)
		self.commerceBox.setGeometry(commerceBox_w, box_h, box_dim, box_dim)
		self.commerceLabel = QLabel('CommerceHub', self)
		self.commerceLabel.move(commerceBox_w + 5, label_height)

		self.commerceBox.mouseReleaseEvent = self.loadCommerceHub
		self.commerceLabel.mouseReleaseEvent = self.loadCommerceHub
		
		# Staples
		self.staplesBox = QFrame(self)
		self.staplesBox.setStyleSheet("QWidget { background-color: %s}" %
			grey.name())
		staplesBox_w = (program_w - box_dim - box_offset)
		self.staplesBox.setGeometry(staplesBox_w, box_h, box_dim, box_dim)
		self.staplesLabel = QLabel('Staples', self)
		self.staplesLabel.move(staplesBox_w + 24, label_height)

		self.staplesBox.mouseReleaseEvent = self.loadStaples
		self.staplesLabel.mouseReleaseEvent = self.loadStaples


		processButton = QPushButton("Process", self)
		processButton.move(program_w - 235, program_h - 50)
		processButton.clicked.connect(self.processing)

		exitButton = QPushButton("Exit", self)
		exitButton.move(program_w - 125, program_h - 50)
		exitButton.clicked.connect(self.close)


		self.setGeometry(300, 300, program_w, program_h)
		self.setWindowTitle('Order Opener')
		self.show()

	def loadGroupon(self, event):

		fname = QFileDialog.getOpenFileName(self, 'Open file', '/')

		if fname[0]:
			var_path = fname[0]
			self.sheetCheck = True
			self.grouponPath = os.path.abspath(var_path)
			self.statusBar().showMessage('Groupon added!')
			self.grouponBox.setStyleSheet("QWidget { background-color: %s}" %
			self.dark_grey.name())


	def loadCommerceHub(self, event):

		fname = QFileDialog.getOpenFileName(self, 'Open file', '/')

		if fname[0]:
			var_path = fname[0]
			self.sheetCheck = True
			self.commercehubPath = os.path.abspath(var_path)
			self.statusBar().showMessage('CommerceHub added!')
			self.commerceBox.setStyleSheet("QWidget { background-color: %s}" %
			self.dark_grey.name())

	def loadStaples(self, event):

		fname = QFileDialog.getOpenFileName(self, 'Open file', '/')

		if fname[0]:
			var_path = fname[0]
			self.sheetCheck = True
			self.staplesPath = os.path.abspath(var_path)
			self.statusBar().showMessage('Staples added!')
			self.staplesBox.setStyleSheet("QWidget { background-color: %s}" %
			self.dark_grey.name())

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

