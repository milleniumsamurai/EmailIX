#-------------------------------------------------------------------------------
# Name:        EmailIX
# Purpose: Pulls email addresses from all pdf files in a directory. Initially
#          conceived as a tool for resume handling and the first piece of an
#          Applicant Tracking System
#
# Author:      Arly Parent
#
# Created:     30/07/2017
# Copyright:   (c) Arly Parent 2017
#-------------------------------------------------------------------------------

from PySide import QtCore, QtGui
from PySide.QtCore import *
from PySide.QtGui import *
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from cStringIO import StringIO
import re


class EmailIX(QtGui.QDialog):
    def __init__(self, parent=None):
        super(EmailIX, self).__init__(parent)

        self.browseButton = self.createButton("&Browse...", self.browse)
        self.findButton = self.createButton("&List Emails", self.find)
##        self.getEmailsButton = self.createButton("&List Emails", self.listEmails(foundFiles))

        self.fileComboBox = self.createComboBox("*.pdf")
        self.textComboBox = self.createComboBox()
        self.directoryComboBox = self.createComboBox(QtCore.QDir.currentPath())
        self.textEditor = QTextBrowser() #This is a multi-line plain-text editor

        fileLabel = QtGui.QLabel("Named:")
        textLabel = QtGui.QLabel("Containing text:")
        directoryLabel = QtGui.QLabel("In directory:")
        self.filesFoundLabel = QtGui.QLabel()

        self.createFilesTable()

        buttonsLayout = QtGui.QHBoxLayout()
        buttonsLayout.addStretch()
        buttonsLayout.addWidget(self.findButton)
##        buttonsLayout.addWidget(self.getEmailButton)

        mainLayout = QtGui.QGridLayout()
        mainLayout.addWidget(fileLabel, 0, 0)
        mainLayout.addWidget(self.fileComboBox, 0, 1, 1, 2)
        mainLayout.addWidget(textLabel, 1, 0)
        mainLayout.addWidget(self.textComboBox, 1, 1, 1, 2)
        mainLayout.addWidget(directoryLabel, 2, 0)
        mainLayout.addWidget(self.directoryComboBox, 2, 1)
        mainLayout.addWidget(self.browseButton, 2, 2)
        mainLayout.addWidget(self.filesTable, 3, 0, 1, 6)
        mainLayout.addWidget(self.filesFoundLabel, 4, 0)
        mainLayout.addWidget(self.textEditor,3, 7, 1, 1)
        mainLayout.addLayout(buttonsLayout, 5, 0, 1, 3)
        self.setLayout(mainLayout)

        self.setWindowTitle("EmailIX - Basic ATS")
        self.resize(1000, 600)

    def browse(self):
        directory = QtGui.QFileDialog.getExistingDirectory(self, "Find Files",
                QtCore.QDir.currentPath())

        if directory:
            if self.directoryComboBox.findText(directory) == -1:
                self.directoryComboBox.addItem(directory)

            self.directoryComboBox.setCurrentIndex(self.directoryComboBox.findText(directory))

    @staticmethod
    def updateComboBox(comboBox):
        if comboBox.findText(comboBox.currentText()) == -1:
            comboBox.addItem(comboBox.currentText())
#------------------------------------------------------------------------------------------------------------------------------------------------------#
    def updateTextBox(self,files):

        pathDir=[]
        fileNameList = files
        for k in files:
            pathDir.append(self.currentDir.absoluteFilePath(k))

        emailList=[]
        for k,index in zip(pathDir,range(len(pathDir))):
            emailList.insert(index,self.getEmails(k))
            self.textEditor.append("%s" % emailList[index])


#------------------------------------------------------------------------------------------------------------------------------------------------------#
    def getEmails(self,path):
        pathlist = path
        rsrcmgr = PDFResourceManager()
        retstr = StringIO()
        codec = 'utf-8'
        laparams = LAParams()
        device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
        fp = file(path, 'rb')
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        password = ""
        maxpages = 0
        caching = True
        pagenos=set()


        for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
            interpreter.process_page(page)

        text = retstr.getvalue()
        email_match = re.search("[\w\.-]+@[\w\.-]+\.\w+",text)
        if email_match:
            email = str(email_match.group(0))
            name_matcher = email.split("@")[0]

        else:
            email = "No email found"
            name_matcher ="No name found"
        result = [name_matcher,email]
        #gotta figure out a way to pull the name from the doc.
        #1) 2 words separated by a space
        #2) Each word starts with a capital letter ^A-Z
        #3) Matches some part of the email string?
        fp.close()
        device.close()
        retstr.close()
        return email
#-----------------------------------------------------------------------------------------------------------------------------------------------------#
    def find(self):
        self.filesTable.setRowCount(0)

        fileName = self.fileComboBox.currentText()
        text = self.textComboBox.currentText()
        path = self.directoryComboBox.currentText()

        self.updateComboBox(self.fileComboBox)
        self.updateComboBox(self.textComboBox)
        self.updateComboBox(self.directoryComboBox)

        self.currentDir = QtCore.QDir(path)
        if not fileName:
            fileName = "*.pdf"
        files = self.currentDir.entryList([fileName],
                QtCore.QDir.Files | QtCore.QDir.NoSymLinks)

        if text:
            files = self.findFiles(files, text)
        self.showFiles(files)


        self.updateTextBox(files)

#------------------------------------------------------------------------------------------------------------------------------------------------------#
    def findFiles(self, files, text):
        progressDialog = QtGui.QProgressDialog(self)

        progressDialog.setCancelButtonText("&Cancel")
        progressDialog.setRange(0, files.count())
        progressDialog.setWindowTitle("Find Files")

        global foundFiles
        foundFiles = []

        for i in range(files.count()):
            progressDialog.setValue(i)
            progressDialog.setLabelText("Searching file number %d of %d..." % (i, files.count()))
            QtGui.qApp.processEvents()

            if progressDialog.wasCanceled():
                break

            inFile = QtCore.QFile(self.currentDir.absoluteFilePath(files[i]))

            if inFile.open(QtCore.QIODevice.ReadOnly):
                stream = QtCore.QTextStream(inFile)
                while not stream.atEnd():
                    if progressDialog.wasCanceled():
                        break
                    line = stream.readLine()
                    if text in line:
                        foundFiles.append(files[i])
                        break

        progressDialog.close()

        return foundFiles
#------------------------------------------------------------------------------------------------------------------------------------------------------#
    def showFiles(self, files):
        for fn in files:
            file_ = QtCore.QFile(self.currentDir.absoluteFilePath(fn))
            size = QtCore.QFileInfo(file_).size()
            dateCreated = QDateTime.toString(QtCore.QFileInfo(file_).created())
            dateCreated_ = str(dateCreated)
            dateCreated_.encode("ascii","ignore")


            fileNameItem = QtGui.QTableWidgetItem(fn)
            fileNameItem.setFlags(fileNameItem.flags() ^ QtCore.Qt.ItemIsEditable)
            sizeItem = QtGui.QTableWidgetItem("%d KB" % (int((size + 1023) / 1024)))
            sizeItem.setTextAlignment(QtCore.Qt.AlignVCenter | QtCore.Qt.AlignRight)
            sizeItem.setFlags(sizeItem.flags() ^ QtCore.Qt.ItemIsEditable)
            createdDateItem = QtGui.QTableWidgetItem("%s" % dateCreated)
            createdDateItem.setTextAlignment(QtCore.Qt.AlignVCenter | QtCore.Qt.AlignRight)
            createdDateItem.setFlags(sizeItem.flags() ^ QtCore.Qt.ItemIsEditable)

            row = self.filesTable.rowCount()
            self.filesTable.insertRow(row)
            self.filesTable.setItem(row, 0, fileNameItem)
            self.filesTable.setItem(row, 1, createdDateItem)
            self.filesTable.setItem(row, 2, sizeItem)

        self.filesFoundLabel.setText("%d file(s) found (Double click on a file to open it)" % len(files))
#------------------------------------------------------------------------------------------------------------------------------------------------------#
    def createButton(self, text, member):
        button = QtGui.QPushButton(text)
        button.clicked.connect(member)
        return button
#------------------------------------------------------------------------------------------------------------------------------------------------------#
    def createComboBox(self, text=""):
        comboBox = QtGui.QComboBox()
        comboBox.setEditable(True)
        comboBox.addItem(text)
        comboBox.setSizePolicy(QtGui.QSizePolicy.Expanding,
                QtGui.QSizePolicy.Preferred)
        return comboBox
#------------------------------------------------------------------------------------------------------------------------------------------------------#
    def createFilesTable(self):
        self.filesTable = QtGui.QTableWidget(0, 3)
        self.filesTable.setSelectionBehavior(QtGui.QAbstractItemView.SelectRows)

        self.filesTable.setHorizontalHeaderLabels(("File Name", "Created","Size"))
        self.filesTable.horizontalHeader().setResizeMode(0, QtGui.QHeaderView.Stretch)
        self.filesTable.verticalHeader().hide()
        self.filesTable.setShowGrid(False)

        self.filesTable.cellActivated.connect(self.openFileOfItem)
#------------------------------------------------------------------------------------------------------------------------------------------------------#
    def openFileOfItem(self, row, column):
        item = self.filesTable.item(row, 0)

        QtGui.QDesktopServices.openUrl(QtCore.QUrl(self.currentDir.absoluteFilePath(item.text())))
#------------------------------------------------------------------------------------------------------------------------------------------------------#


if __name__ == '__main__':

    import sys


    app = QtGui.QApplication(sys.argv)
    emailIX = EmailIX()
    emailIX.show()
    sys.exit(app.exec_())