#!/opt/local/bin/python

import os 
import sys
import shutil
import subprocess
import types
from PyQt4.QtCore import * 
from PyQt4.QtGui import * 
import docx2xelatex

class ZinkLoadTab(QWidget):
	def __init__(self, mainwindow):
		QWidget.__init__(self)
		self.mw = mainwindow
		self.setupTab()
		
	def setupTab(self):
		maingrid = QVBoxLayout()
				
		doclabelgrid = QHBoxLayout()
		doclabel = QLabel('Current document:')
		doclabelgrid.addWidget(doclabel)
		self.mw.docxfilenamelabel = QLabel('None')
		doclabelgrid.addWidget(self.mw.docxfilenamelabel)
		loadbutton = QPushButton('Load file...')
		self.connect(loadbutton, SIGNAL("clicked()"), self.openFile)
		doclabelgrid.addWidget(loadbutton)
		maingrid.addLayout(doclabelgrid)
		
		creatorgrid = QHBoxLayout()
		creatorlabel = QLabel("Author:")
		creatorlabel.setMinimumWidth(75)
		creatorgrid.addWidget(creatorlabel)
		self.mw.creatorentry = QLineEdit()
		creatorgrid.addWidget(self.mw.creatorentry)
		#creatorgrid.addStretch(1)
		maingrid.addLayout(creatorgrid)
		
		titlegrid = QHBoxLayout()
		titlelabel = QLabel("Title:")
		titlelabel.setMinimumWidth(75)
		titlegrid.addWidget(titlelabel)
		self.mw.titleentry = QLineEdit()
		titlegrid.addWidget(self.mw.titleentry)
		#titlegrid.addStretch(1)
		maingrid.addLayout(titlegrid)
		
		subtitlegrid = QHBoxLayout()
		subtitlelabel = QLabel("Subtitle:")
		subtitlelabel.setMinimumWidth(75)
		subtitlegrid.addWidget(subtitlelabel)
		self.mw.subtitleentry = QLineEdit()
		subtitlegrid.addWidget(self.mw.subtitleentry)
		#subtitlegrid.addStretch(1)
		maingrid.addLayout(subtitlegrid)
		
		maingrid.addWidget(QLabel(""))
		
		latexsettingsgroup = QGroupBox("LaTeX settings")
		
		latexsettingslayout = QVBoxLayout()
		
		dstylinggrid = QHBoxLayout()
		dstylingcbox = QCheckBox("Use direct styling")
		dstylingcbox.setChecked(True)
		dstylinggrid.addWidget(dstylingcbox)
		dstylinggrid.addStretch(1)
		latexsettingslayout.addLayout(dstylinggrid)
		
		tcolorgrid = QHBoxLayout()
		tcolorcbox = QCheckBox("Use table colors")
		tcolorgrid.addWidget(tcolorcbox)
		tcolorgrid.addStretch(1)
		latexsettingslayout.addLayout(tcolorgrid)
		
		latexsettingsgroup.setLayout(latexsettingslayout)
		
		
		maingrid.addWidget(latexsettingsgroup)
		maingrid.addStretch(1)
		self.setLayout(maingrid)
		
		# settings
		# - table colors
		# - direct styling
		# - zink style
		# - point size
		# - paper size
		# - table suppression
		# - discard frontmatter
		# - titlepage selection
		# - author info
		# - Title/subtitle
		# - style translation:
		#   - headings
		#   - floats
		#   - captions
		#   - environments
		#   - direct styling commands 
		
	def openFile(self):
		filename = str(QFileDialog.getOpenFileName(self, "Choose a docx file", ".", "*.docx *.docm"))
		if filename:
			self.mw.docxfilename = filename
			self.mw.docxfilenamelabel.setText(os.path.split(self.mw.docxfilename)[1])
			self.mw.docx = docx2xelatex.DocX(filename)
			self.mw.creatorentry.setText(self.mw.docx.getCreator())
			self.mw.D2X = docx2xelatex.docx2xelatex(self.mw.docx)
		

class MainWindow(QMainWindow):
	def __init__(self, parent=None):
		super(MainWindow, self).__init__(parent)
		self.docxfilename = ""

		# GUI
		self.tabs = QTabWidget()
		self.tabLoad = ZinkLoadTab(self) # docx loading and conversion settings
		
		#self.tabReferences = ZinkReferencesTab(self) # bookmarks and references
		#self.tabFigures = ZinkFiguresTab(self)
		#self.tabTables = ZinkTablesTab(self)
		
		#self.tabSettings = ZinkSettingsTab(self) # LaTeX document settings
		#self.tabProcess = ZinkProcessTab(self) # Actual conversion
	
		self.tabs.addTab(self.tabLoad, 'Load file')
		
		self.setCentralWidget(self.tabs)
		
		
	
def main(): 
	app = QApplication(sys.argv)
	app.setOrganizationName("Zink Typografie")
	app.setOrganizationDomain("zinktypografie.nl")
	app.setApplicationName("Zink docx to XeLaTex converter")
	#app.setWindowIcon(QIcon(":/icon.png"))
	zmw = MainWindow()
	zmw.show()
	zmw.raise_()
	app.exec_()

main()