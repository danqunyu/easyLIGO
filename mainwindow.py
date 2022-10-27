from PyQt5 import QtCore, QtGui, QtWidgets
from easyLIGO import Ui_MainWindow
from PyQt5.QtWidgets import (QMainWindow, QApplication, QWidget, QPushButton, QHBoxLayout, QLineEdit, QMessageBox,
                             QFileDialog)
import pandas as pd


class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)

        # 点击“打开文件”，显示对应文件名
        self.pb_OpenConfigFile.clicked.connect(self.OpenFileAndConvert)

    def OpenFileAndConvert(self):
        # 打开文件并确定输入/输出文件名
        inputFile, fileType = QFileDialog.getOpenFileName(parent=None, caption='Open Config File', directory='',
                                                          filter='Excel files(*.xlsx , *.xls)')
        self.tx_OpenFile.setText(inputFile)
        self.PrintRecord("Successfully read file: %s" % inputFile)

        filepath = "/".join(inputFile.split('/')[:-1])
        outputFile = self.GetOutFile(filepath)

        # config文件转换为UTC所需condition table
        self.CovertConfigFile(inputFile, outputFile)

    def GetOutFile(self, filepath):
        filename = self.le_OutFileName.text()
        if len(filename) == 0:
            QMessageBox.critical(self, 'Empty FileName or FilePath', 'Please set your destination file')
        else:
            rst = (filepath + '/' + filename)
            self.PrintRecord("Successfully set output file: %s" % rst)
            return rst
            # raise Exception("Empty FileName or FilePath")

    def CovertConfigFile(self, filename, conditionTable):
        # 读取文件，并校验
        dfDic = pd.read_excel(filename, sheet_name=None)  # 检查config excel中只有两个sheet，且名字正确
        for kvp in dfDic.items():
            if kvp[0] not in ['Config', 'MethodLib']:
                self.PrintRecord("Warning: sheet %s is not in file" % kvp[0])
                QMessageBox.critical(self, 'Warning', 'Redundant sheets except "Config and MethodLib" in input file')
            else:
                self.PrintRecord("Sheet \"%s\" is in file" % kvp[0])

        config = dfDic['Config']
        methodLib = dfDic['MethodLib']

        # 生成TestInstance表格
        TestInstance = pd.merge(config[['TestInstanceName', 'TestMethod']], methodLib, on='TestMethod', how='left')
        TestInstance.where(TestInstance.notnull(), '')  # 将NaN替换为空
        TestInstance.to_excel(conditionTable, 'TestInstance', index=False)  # 新建excel
        self.PrintRecord("Sheet \"TestInstance\" is exported to condition table.")

        # 生成MainFlow和DynamicPreload(如果有的话)表格
        writer = pd.ExcelWriter(conditionTable, mode='a', engine='openpyxl')  # 通过excelWriter新增sheet
        for item in config['TestFlowName'].drop_duplicates().values:
            tempDf = config.iloc[list(config['TestFlowName'] == item)].loc[:, ['TestSuiteName', 'TestInstanceName']]
            tempDf['Opcode'] = 'test'
            tempDf['FlowNodeSettings'] = 'ContinueOnFail'
            tempDf.rename({'TestInstanceName': 'Parameter'}, axis='columns', inplace=True)
            output = tempDf[['Opcode', 'Parameter', 'TestSuiteName', 'FlowNodeSettings']]
            output.to_excel(writer, item, index=False)
            self.PrintRecord("Sheet \"%s\" is exported to condition table." % item)

        # 生成LMT表格

        # 结束Excel Writer进程
        writer.close()

    # 输出日志记录，在tx_OperationRecord中不断追加
    def PrintRecord(self, item):
        self.tx_OperationRecord.append(item)
        QApplication.processEvents()
        pass
