from PyQt5.QtWidgets import QApplication, QListWidget, QMainWindow, QPushButton, QLabel, QVBoxLayout, QWidget, \
    QProgressBar, QDialog, QLineEdit, QMenu, QAction, QTableWidget, QTableWidgetItem, QToolButton, QFileDialog
from PyQt5.QtCore import QEvent
from PyQt5 import QtCore, QtGui, QtWidgets
from pathlib import Path
import sys
import win32com.client as win32
import os
import pandas as pd
import openpyxl
import xlsxwriter
import webbrowser
import asyncio

from random import randint
from functools import partial

listExportOrders = []
listExportFiles = []
global errorWidget


def openFolder(path):
    if path == Path.cwd() / "output":
        createFolder = AutoEmail()
        os.path.realpath(createFolder.createFolder(path))
    os.startfile(path)


def openExel():
    os.system("start EXCEL.EXE validateM.xlsm")


def DownloadPDFs(progressWid, errorWid, outlookFolder, path):
    start = AutoEmail()
    start.init(progressWid, errorWid, outlookFolder, path)


def validatePDFs(tableMissingFiles, nMissingFiles, l_totalFiles, progressWindow, path):
    progressWindow.ValidateProgress("PDFs", 0, 1)
    tableMissingFiles.clear()
    tableMissingFiles.setHorizontalHeaderLabels(["Order", "Destino", "Transportador", "Data"])
    listExportFiles.clear()
    isExist = os.path.exists(path)

    if isExist:
        filesName = getFilesNamesinFolder(path)
        OrderData = getExelData()
        row = 1
        missing = 0

        for order in OrderData.Order:
            found = False
            row = row + 1
            for filename in filesName:
                fileName = os.path.splitext(filename)[0]
                splitFileName = fileName.split("_")
                for spl in splitFileName:
                    if str(order) == spl:
                        found = True
                        break
            if not found:
                x = OrderData[OrderData['Order'] == order]
                strDestino = str(x.DESTINO).replace("Name: DESTINO, dtype: object", "")
                strTransportador = str(x.TRANSPORTADOR).replace("Name: TRANSPORTADOR, dtype: object", "")
                strData = str(x.DATA).replace("Name: DATA, dtype: datetime64[ns]", "")
                splitDestino = strDestino.split("   ")
                splitTransportador = strTransportador.split("   ")
                splitData = strData.split("   ")
                if len(splitDestino) > 1:
                    missing = missing + 1
                    tableMissingFiles.setRowCount(missing)
                    tableMissingFiles.setItem(missing - 1, 0, QTableWidgetItem(str(order)))
                    tableMissingFiles.setItem(missing - 1, 1, QTableWidgetItem(splitDestino[1]))
                    tableMissingFiles.setItem(missing - 1, 2, QTableWidgetItem(splitTransportador[1]))
                    tableMissingFiles.setItem(missing - 1, 3, QTableWidgetItem(splitData[1]))
                    listExportFiles.append(
                        str(row) + " | " + str(order) + " - " + splitDestino[1] + " - " + splitTransportador[
                            1] + " - " + splitData[1])
            progressWindow.ValidateProgress("PDFs", row, OrderData.__len__())

        nMissingFiles.setText(str(missing))
        nMissingFiles.adjustSize()
        l_totalFiles.setText("Total Files: " + str(filesName.__len__()))
        l_totalFiles.adjustSize()
        tableMissingFiles.resizeColumnsToContents()
        progressWindow.SelfClose()


def validateOrders(listMissingOrders, nMissingOrders, l_totalOrders, progressWindow, path):
    progressWindow.ValidateProgress("Pedidos", 0, 1)
    listMissingOrders.clear()
    listExportOrders.clear()
    orderDict = {}

    isExist = os.path.exists(path)

    if isExist:
        filesName = getFilesNamesinFolder(path)
        OrderData = getExelData()
        row = 0
        missing = 0
        rowCode = 0

        for filename in filesName:
            row = row + 1
            fileName = os.path.splitext(filename)[0]
            splitFileName = fileName.split("_")
            if splitFileName.__len__() >= 2:
                rowCode = rowCode + 1
            for spl in splitFileName:
                found = False
                if spl == "PV" or spl == "ECI":
                    break
                for order in OrderData.Order:
                    if str(order) == spl:
                        found = True
                        break

                if not found:
                    orderDict.update({spl: fileName})
                    twoLines = ""
                    if splitFileName.__len__() >= 2:
                        twoLines = "   <-- " + str(rowCode)
                    strtes = str(row) + " | " + spl + ".pdf" + twoLines
                    listMissingOrders.addItem(strtes)
                    missing = missing + 1
                    listExportOrders.append(str(row) + " | " + spl + ".pdf")
            progressWindow.ValidateProgress("Pedidos", row, filesName.__len__())
        nMissingOrders.setText(str(missing))
        nMissingOrders.adjustSize()
        l_totalOrders.setText("Total Orders: " + str(OrderData.__len__()))
        l_totalOrders.adjustSize()
        progressWindow.SelfClose()
        return orderDict


def getFilesNamesinFolder(dir):
    return os.listdir(dir)


def getExelData():
    data = pd.read_excel('validateM.xlsm')
    return data


def ExportToExcel():
    CreateExportExcel()
    os.system("start EXCEL.EXE Export.xlsx")

    workbook = xlsxwriter.Workbook('Export.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.write(0, 0, "Missing Files")
    worksheet.write(0, 10, "Missing Orders")

    for row_index, row in enumerate(listExportFiles):
        worksheet.write(row_index + 1, 0, row)

    for row_index, row in enumerate(listExportOrders):
        worksheet.write(row_index + 1, 10, row)

    workbook.close()


def CreateExportExcel():
    if not os.path.exists("Export.xlsx"):
        path = Path.cwd() / "Export.xlsx"
        wb = openpyxl.Workbook()
        wb.save(path)


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.missingOrderDict = None
        self.setFixedSize(1000, 650)
        self.setWindowTitle("Auto Email")

        self.outlookFolder = ""
        self.path = ""

        menuBar = self.menuBar()
        editmenu = QMenu("&Edit", self)
        menuBar.addMenu(editmenu)
        self.changeOutlookFolder = QAction("Mudar Pasta do Outlook")
        self.changeOutlookFolder.triggered.connect(self.ChangeOutlookFolder)
        self.reset = QAction("Resetar Configs")
        self.reset.triggered.connect(self.Reset)
        editmenu.addAction(self.changeOutlookFolder)
        editmenu.addAction(self.reset)

        self.l_outlookFolder = QLabel("", self)
        self.l_outlookFolder.move(50, 30)
        self.l_outlookFolder.setStyleSheet('font-size:20px')
        self.l_outlookFolder.adjustSize()

        btnSelectFolder = QPushButton("Selecionar Pasta do Outlook", self)
        btnSelectFolder.setGeometry(400, 30, 150, 30)
        btnSelectFolder.clicked.connect(self.SelectOutlookFolder)

        btnDownloadPDFs = QPushButton("Download PDFs", self)
        btnDownloadPDFs.setGeometry(50, 120, 200, 50)
        btnDownloadPDFs.clicked.connect(self.DownloadPDFs)

        self.nMissingFiles = QLabel("0", self)
        self.nMissingFiles.move(210, 200)
        self.nMissingFiles.setStyleSheet('font-size:30px')
        self.nMissingFiles.adjustSize()

        self.tableMissingFiles = QTableWidget(self)
        self.tableMissingFiles.setGeometry(50, 250, 513, 300)
        self.tableMissingFiles.setColumnCount(4)
        self.tableMissingFiles.setHorizontalHeaderLabels(["Order", "Destino", "Transportador", "Data"])

        self.nMissingOrders = QLabel("0", self)
        self.nMissingOrders.move(800, 200)
        self.nMissingOrders.setStyleSheet('font-size:30px')
        self.nMissingOrders.adjustSize()

        self.l_totalFiles = QLabel("Total Files: ", self)
        self.l_totalFiles.move(50, 550)
        self.l_totalFiles.setStyleSheet('font-size:14px')
        self.l_totalFiles.adjustSize()

        self.l_totalOrders = QLabel("Total Orders: ", self)
        self.l_totalOrders.move(600, 550)
        self.l_totalOrders.setStyleSheet('font-size:14px')
        self.l_totalOrders.adjustSize()

        self.listMissingOrders = QListWidget(self)
        self.listMissingOrders.setGeometry(600, 250, 375, 300)
        self.listMissingOrders.installEventFilter(self)
        self.listMissingOrders.doubleClicked.connect(self.OpenPDF)

        btnVerifyPDFs = QPushButton("Verificar PDFs", self)
        btnVerifyPDFs.setGeometry(260, 120, 200, 50)
        btnVerifyPDFs.clicked.connect(self.validate)

        btnOpenPDFsFolder = QPushButton("Abrir Pasta dos PDFs", self)
        btnOpenPDFsFolder.setGeometry(580, 120, 200, 50)
        btnOpenPDFsFolder.clicked.connect(self.OpenPDFsFolder)

        btnOpenExcelFile = QPushButton("Abrir ficheiro Excel", self)
        btnOpenExcelFile.setGeometry(790, 120, 200, 50)
        btnOpenExcelFile.clicked.connect(openExel)

        missingFiles = QLabel("Falta PDFs:", self)
        missingFiles.move(50, 200)
        missingFiles.setStyleSheet('font-size:30px')
        missingFiles.adjustSize()

        missingOrders = QLabel("Falta no Ecxel:", self)
        missingOrders.move(600, 200)
        missingOrders.setStyleSheet('font-size:30px')
        missingOrders.adjustSize()

        btnOpenExcelFile = QPushButton("Exportar para ficheiro Excel", self)
        btnOpenExcelFile.setGeometry(50, 580, 200, 50)
        btnOpenExcelFile.clicked.connect(ExportToExcel)

        self.toolButtonOpenDialog = QtWidgets.QToolButton(self)
        self.toolButtonOpenDialog.setGeometry(QtCore.QRect(50, 70, 120, 25))
        self.toolButtonOpenDialog.setText("Mudar Pasta destino")
        self.toolButtonOpenDialog.clicked.connect(self._open_file_dialog)

        self.l_directory = QLabel("", self)
        self.l_directory.move(175, 73)
        self.l_directory.setStyleSheet('font-size:14px')
        self.l_directory.adjustSize()

        self.dialog = SecWindow()
        self.selectOutlookFolder = SelectOutlookFolder()
        self.error = ErrorMessage()
        self.progress = ProgressWindow()
        self.confirmWidow = ConfrimWindow()
        self.Updatelabels()

    def Reset(self):
        path = Path.cwd() / "config.txt"
        try:
            os.remove(path)
        except:
            pass
        self.Updatelabels()

    def OpenPDFsFolder(self):
        openFolder(self.path)

    def _open_file_dialog(self):
        directory = str(QtWidgets.QFileDialog.getExistingDirectory())
        if directory != "":
            self.UpdatefileConfig("Directory", directory)
            self.Updatelabels()

    def UpdatefileConfig(self, configToChange, newConfig):
        try:
            r_file = open("config.txt", "r")
            lines = r_file.readlines()
            r_file.close()
            configFound = False

            w_file = open("config.txt", "w")
            for i, line in enumerate(lines):
                config = line.split(": ")
                if config[0] == configToChange:
                    newLine = configToChange + ": " + newConfig
                    w_file.writelines(str(newLine) + "\n")
                    configFound = True
                else:
                    w_file.writelines(line)
            w_file.close()

            if not configFound:
                w_file = open("config.txt", "a")
                newLine = configToChange + ": " + newConfig
                w_file.writelines(str(newLine) + "\n")
                w_file.close()
        except:
            w_file = open("config.txt", "w")
            w_file.write(configToChange + ": " + newConfig + "\n")
            w_file.close()

    def Updatelabels(self):
        try:
            config = open('config.txt', 'r')
            lines = config.readlines()
            configDirectoryFound = False
            configOutlookFolderFound = False

            for i, line in enumerate(lines):
                config = line.split(": ")
                if config[0] == "Directory":
                    self.l_directory.setText(config[1])
                    self.l_directory.adjustSize()
                    self.path = config[1].strip()
                    configDirectoryFound = True
                if config[0] == "Outlook Folder":
                    self.l_outlookFolder.setText("Outlook Folder: " + config[1])
                    self.l_outlookFolder.adjustSize()
                    self.outlookFolder = config[1].strip()
                    configOutlookFolderFound = True

            if not configDirectoryFound:
                self.path = Path.cwd() / "output"
                self.l_directory.setText(str(self.path))
                self.l_directory.adjustSize()

            if not configOutlookFolderFound:
                self.l_outlookFolder.setText("Outlook Folder: testeAuto")
                self.l_outlookFolder.adjustSize()
                self.outlookFolder = "testeAuto"
        except:
            self.l_outlookFolder.setText("Outlook Folder: testeAuto")
            self.l_outlookFolder.adjustSize()
            self.outlookFolder = "testeAuto"
            self.path = Path.cwd() / "output"
            self.l_directory.setText(str(self.path))
            self.l_directory.adjustSize()

    def ChangeOutlookFolder(self):
        self.dialog.show()
        self.dialog.ChangeOutlookFolder(self)

    def SelectOutlookFolder(self):
        self.selectOutlookFolder.show()
        self.selectOutlookFolder.FillOutlookFoldersList(self, self.error)

    def DownloadPDFs(self):
        DownloadPDFs(self.progress, self.error, self.outlookFolder, self.path)

    def validate(self):
        self.progress.show()
        validatePDFs(self.tableMissingFiles, self.nMissingFiles, self.l_totalFiles, self.progress, self.path)
        self.progress.show()
        self.missingOrderDict = validateOrders(self.listMissingOrders, self.nMissingOrders, self.l_totalOrders,
                                               self.progress, self.path)

    def getMultiRowByRowCode(self, rowCode):
        itens = []
        for x in range(self.listMissingOrders.count()):
            aux = self.listMissingOrders.item(x).text().split("<-- ")
            if aux.__len__() > 1:
                if aux[1] == rowCode:
                    aux1 = self.listMissingOrders.item(x).text().split(" | ")
                    aux2 = aux1[1].split(".")
                    itens.append(str(aux2[0]))
        return itens

    def ValidateNewOrderName(self, newOder, oldName):
        teste = oldName.split("_")
        found = False
        foundCount = 0
        for oldOrder in teste:
            if newOder == oldOrder:
                found = True
                foundCount = foundCount + 1

        if not found:
            FileNames = getFilesNamesinFolder(self.path)
            for file in FileNames:
                file = file.split(".")
                spl = file[0].split("_")
                for order in spl:
                    if order == newOder:
                        return True
        return False

    def ChangeFileName(self):
        try:
            fileName = self.listMissingOrders.currentItem().text()
            fileName = fileName.split("| ")
            order = fileName[1].split(".")
            editFileName = self.missingOrderDict[order[0]]
            self.dialog.show()
            self.dialog.getData(self.listMissingOrders.currentItem().text(), editFileName, self, self.path)
        except:
            pass

    def DeleteFile(self):
        try:
            fileName = self.listMissingOrders.currentItem().text()
            fileName = fileName.split("| ")
            order = fileName[1].split(".")
            editFileName = self.missingOrderDict[order[0]]
            self.confirmWidow.SelfShow(editFileName, self)
        except:
            pass

    def OpenPDF(self):
        fileName = self.listMissingOrders.currentItem().text()
        fileName = fileName.split("| ")
        order = fileName[1].split(".")
        editFileName = self.missingOrderDict[order[0]]

        path = os.path.join(self.path, editFileName + ".pdf")
        os.startfile(path)

    def getSelectedRow(self):
        return listMissingOrders.currentItem().text()

    def eventFilter(self, source, event):
        if event.type() == QEvent.ContextMenu and source is self.listMissingOrders:
            buttonEdit_Action = QAction("Editar", self)
            buttonEdit_Action.setStatusTip("Mudar Nome")
            buttonEdit_Action.triggered.connect(self.ChangeFileName)

            buttonDelete_Action = QAction("Apagar", self)
            buttonDelete_Action.setStatusTip("Apagar PDF")
            buttonDelete_Action.triggered.connect(self.DeleteFile)

            menu = QMenu()
            menu.addAction(buttonEdit_Action)
            menu.addAction(buttonDelete_Action)

            menu.exec_(event.globalPos())
            return True
        return super().eventFilter(source, event)


class SecWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.path = None
        self.main = None
        self.setFixedSize(300, 120)
        self.setWindowTitle("Auto Email")

        self.pdfName = ""
        self.operation = ""

        self.l_operation = QLabel("", self)
        self.l_operation.move(15, 5)
        self.l_operation.setStyleSheet('font-size:14px')
        self.l_operation.adjustSize()

        self.l_pdfPosition = QLabel("", self)
        self.l_pdfPosition.move(7, 28)
        self.l_pdfPosition.setStyleSheet('font-size:14px')
        self.l_pdfPosition.adjustSize()

        self.i_input = QLineEdit(self)
        self.i_input.setGeometry(25, 25, 230, 20)

        self.l_pdfext = QLabel("", self)
        self.l_pdfext.move(255, 30)
        self.l_pdfext.setStyleSheet('font-size:14px')
        self.l_pdfext.adjustSize()

        btnOpenExcelFile = QPushButton("Save", self)
        btnOpenExcelFile.setGeometry(200, 60, 75, 50)
        btnOpenExcelFile.clicked.connect(self.Save)

        self.errorMessage = ErrorMessage()

    def ChangeOutlookFolder(self, mainWindow):
        self.operation = "changeOutlook"
        self.main = mainWindow
        self.l_operation.setText("Nome da nova Pasta do Outlook")
        self.l_operation.adjustSize()
        self.l_pdfext.setText("")
        self.l_pdfext.adjustSize()
        self.l_pdfPosition.setText("")
        self.l_pdfext.adjustSize()
        self.i_input.clear()

    def getData(self, Data, fileName, mainWindow, path):
        self.path = path
        self.operation = "changePDF"
        self.l_operation.setText("Novo nome para o PDF")
        self.l_operation.adjustSize()
        self.l_pdfext.setText(".pdf")
        self.l_pdfext.adjustSize()

        self.main = mainWindow

        splitFileName = Data.split(" | ")
        self.l_pdfPosition.setText(splitFileName[0])
        self.pdfName = fileName
        self.l_pdfPosition.adjustSize()
        self.i_input.setText(fileName)

    def Save(self):
        if self.operation == "changePDF":
            orderName = ""
            twoOrdersSameName = False
            inputText = str(self.i_input.text())
            for order in inputText.split("_"):
                if self.main.ValidateNewOrderName(order, self.pdfName):
                    twoOrdersSameName = True
                    orderName = order

            if not twoOrdersSameName:
                newPDFName = (str(self.i_input.text()) + ".pdf")
                try:
                    os.rename(os.path.join(self.path, self.pdfName + ".pdf"), os.path.join(self.path, newPDFName))
                    self.main.validate()
                    self.close()
                except:
                    self.errorMessage.SelfShow("ERRO!!!: Não foi possivel renomear o arquivo!!", "")
            else:
                self.errorMessage.SelfShow("Ja existe um arquivo com esse numero de Pedido!!", "Numero: " + orderName)

        if self.operation == "changeOutlook":
            newOutlookFolder = (str(self.i_input.text()))

            if len(newOutlookFolder) != 0:
                self.main.UpdatefileConfig("Outlook Folder", newOutlookFolder)
                self.main.Updatelabels()
                self.close()


class SelectOutlookFolder(QMainWindow):

    def __init__(self):
        super().__init__()
        self.outlookFolder = None
        self.main = None
        self.setFixedSize(520, 360)
        self.setWindowTitle("Auto Email")

        self.l_directory = QLabel("Pastas do Outlook", self)
        self.l_directory.move(20, 20)
        self.l_directory.setStyleSheet('font-size:14px')
        self.l_directory.adjustSize()

        self.listOutlookFolders = QListWidget(self)
        self.listOutlookFolders.setGeometry(20, 50, 375, 300)
        self.listOutlookFolders.itemClicked.connect(self.itemClicked_event)

        btnselectOutlookFolder = QPushButton("Selecionar Pasta", self)
        btnselectOutlookFolder.setGeometry(400, 300, 100, 50)
        btnselectOutlookFolder.clicked.connect(self.SelectFolder)

    def itemClicked_event(self):
        selc = self.listOutlookFolders.selectedItems()[0]
        self.outlookFolder = selc.data(0)

    def FillOutlookFoldersList(self, main, errorMessage):
        self.main = main

        try:
            self.listOutlookFolders.clear()
            outlook = win32.Dispatch('Outlook.Application').GetNamespace("MAPI")
            user = outlook.Application.Session.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress
            folders = outlook.Folders[user].Folders

            for folder in folders:
                self.listOutlookFolders.addItem(folder.name)
        except:
            errorMessage.SelfShow("ERRO!!!: Não foi possivel possivel obter as pastas do Outlook!!", "")
            self.close()


    def SelectFolder(self):
        self.main.UpdatefileConfig("Outlook Folder", self.outlookFolder)
        self.main.Updatelabels()
        self.close()


class ProgressWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setFixedSize(300, 60)
        self.setWindowTitle("Auto Email")

        self.l_status = QLabel("", self)
        self.l_status.move(10, 5)
        self.l_status.setStyleSheet('font-size:14px')
        self.l_status.adjustSize()

        self.progressBar = QProgressBar(self)
        self.progressBar.setGeometry(20, 25, 250, 20)

    def DownloadProgress(self, number, total):
        self.l_status.setText("Downloadind: PDFs")
        self.l_status.adjustSize()
        percent = number / total
        self.progressBar.setValue(int(percent * 100))
        QApplication.processEvents()

    def ValidateProgress(self, lStatus, number, total):
        self.l_status.setText("Verificando: " + lStatus)
        self.l_status.adjustSize()
        percent = number / total
        self.progressBar.setValue(int(percent * 100))
        QApplication.processEvents()

    def SelfClose(self):
        self.close()


class ErrorMessage(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setFixedSize(370, 190)
        self.setWindowTitle("ERROR: Auto Email")

        self.l_error = QLabel("Error: ", self)
        self.l_error.move(25, 15)
        self.l_error.setStyleSheet('font-size:20px')
        self.l_error.adjustSize()

        self.l_errormsg = QLabel("", self)
        self.l_errormsg.move(35, 40)
        self.l_errormsg.setStyleSheet('font-size:14px')
        self.l_errormsg.setWordWrap(True)
        self.l_errormsg.adjustSize()

        self.l_errorOrderNumber = QLabel("", self)
        self.l_errorOrderNumber.move(45, 90)
        self.l_errorOrderNumber.setStyleSheet('font-size:12px')
        self.l_errorOrderNumber.setWordWrap(True)
        self.l_errorOrderNumber.adjustSize()

        btnOpenExcelFile = QPushButton("OK", self)
        btnOpenExcelFile.setGeometry(280, 130, 75, 50)
        btnOpenExcelFile.clicked.connect(self.Close)

    def SelfShow(self, errorMSG, orderNumber):
        self.l_errormsg.setText(errorMSG)
        self.l_errormsg.adjustSize()
        if orderNumber != "":
            self.l_errorOrderNumber.setText(str(orderNumber))
        self.l_errorOrderNumber.adjustSize()
        return self.show()

    def Close(self):
        self.close()


class ConfrimWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.orderNumber = None
        self.main = None
        self.setFixedSize(370, 150)
        self.setWindowTitle("AutoEmail")

        self.l_texte = QLabel("Tem a certeza que quer eliminar?", self)
        self.l_texte.move(25, 10)
        self.l_texte.setStyleSheet('font-size:16px')
        self.l_texte.setWordWrap(True)
        self.l_texte.adjustSize()

        self.l_OrderNumber = QLabel("", self)
        self.l_OrderNumber.move(45, 60)
        self.l_OrderNumber.setStyleSheet('font-size:12px')
        self.l_OrderNumber.setWordWrap(True)
        self.l_OrderNumber.adjustSize()

        btnOpenExcelFile = QPushButton("SIM", self)
        btnOpenExcelFile.setGeometry(280, 90, 75, 50)
        btnOpenExcelFile.clicked.connect(self.Delete)

        btnOpenExcelFile = QPushButton("Não", self)
        btnOpenExcelFile.setGeometry(200, 90, 75, 50)
        btnOpenExcelFile.clicked.connect(self.Close)

    def SelfShow(self, orderNumber, main):
        self.l_OrderNumber.adjustSize()
        self.main = main
        self.orderNumber = str(orderNumber) + ".pdf"
        self.l_OrderNumber.setText(self.orderNumber)
        self.l_OrderNumber.adjustSize()

        return self.show()

    def Delete(self):
        path = os.path.join(self.main.path, self.orderNumber)
        os.remove(path)
        self.main.validate()
        self.close()

    def Close(self):
        self.close()


class AutoEmail:

    def init(self, progWindow, errorWid, outlookFolder, newpath):
        if newpath == Path.cwd() / "output":
            self.createFolder(newpath)
        messages = self.connectOutlookGetEmails(errorWid, outlookFolder)
        self.SavePDFs(messages, newpath, progWindow)

    def createFolder(self, path):
        path.mkdir(parents=True, exist_ok=True)
        return path

    def connectOutlookGetEmails(self, errorWid, outlookFolder):
        try:
            outlook = win32.Dispatch('Outlook.Application').GetNamespace("MAPI")
            root_folders = outlook.Folders.Item(1)

            try:
                inbox = root_folders.Folders[outlookFolder]
                return inbox.items
            except Exception as e:
                errorWid.SelfShow("Pasta do Outlook não encontrada!!!", e)
                print(e)
                return ""
        except Exception as e:
            errorWid.SelfShow("Falha a conectar com o Outlook!!!", e)
            print(e)
            return ""

    def SavePDFs(self, messages, path, progWind):
        progWind.show()
        count = 0
        for m in messages:
            attachments = m.Attachments

            for att in attachments:
                newPath = os.path.join(path, att.FileName)
                att.SaveAsFile(newPath)

                count = count + 1
                progWind.DownloadProgress(count, len(messages))
        progWind.SelfClose()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    app.exec()
