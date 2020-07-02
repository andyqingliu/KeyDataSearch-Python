import pandas as pd
import os
import re

excelNameKey = '立项申请表'
moduleColumns = ['文件路径' ,'项目名称', '最终用户', '签约甲方', '商务路线', '项目说明', '申请人（销售）', 
'申请人所属部门', '申请日期', '项目预算', '投标保证金', '联系电话', '投标方式', '投标时间', '付款方式']

class ExcelHandler(object):
    def __init__(self):
        super(ExcelHandler,self).__init__()
        # 明细数据，原始数据
        self.resData = None

        self.files = {}
        self.excelDict = {}
        self.colValuesDict = {}

    def ReadFileFolder(self, fileFolderPath):
        if(len(fileFolderPath) > 0):
            self.files = os.walk(fileFolderPath)
            for root, dirs, files in self.files:
                for filePath in files:
                    absPath = os.path.join(root, filePath)
                    path, fileName = os.path.split(absPath)
                    # fullPath, ext = os.path.splitext(absPath)
                    if excelNameKey in fileName:
                        curExcel = self.ReadFolderExcel(absPath)
                        self.excelDict[absPath] = curExcel

                    # if ext == ".xls":
                    #     curExcel = self.ReadFolderExcel(absPath)
                    #     print("XLS:" + absPath)
                    # elif ext == ".xlsx":
                    #     print("XLSX:" + absPath)

    def ReadFolderExcel(self, fPath):
        if(len(fPath) > 0):
            return pd.read_excel(fPath)

    def SearchContentFromFiles(self):
        if len(moduleColumns) > 0 and len(self.excelDict) > 0:
            for key in self.excelDict.keys():

                value = self.excelDict[key]
                # print(value.columns)
                for col in moduleColumns:
                    listofPos = self.GetIndexes(value, col)
                    finalValue = ""
                    if len(listofPos) > 0:
                        rowIndex = listofPos[0][0]
                        colIndex = listofPos[0][1]
                        finalValue = value.iloc[rowIndex, colIndex + 1]
                    else:
                        if col != "文件路径":
                            print("在文件 ", key, " 没找到的列名：", col)

                    if self.colValuesDict.get(col) == None:
                        print("列名为空，要初始化，列名是：", key, col)
                        colValues = list()
                        if col == "文件路径":
                            finalValue = key

                        colValues.append(finalValue)
                        self.colValuesDict[col] = colValues
                    else:
                        if col == "文件路径":
                            finalValue = key
                        self.colValuesDict[col].append(finalValue)

    def MatchRemoveSpace(self, x, value):
        if isinstance(x, str):
            pattern = re.compile(r'\s+')
            # 通过正则表达式匹配去除文本里面的空格，提高模糊匹配能力(不能简单只去除首尾空格)
            # x = x.strip()
            x = re.sub(pattern, '', x)
            return value in x
        else:
            return False

    def BuildDataFrame(self):
        # for col, colValues in self.colValuesDict.items():
        #     for values in colValues:
        #         print("列名：", col, "列的值：", values)
        #     print("************************")

        df = pd.DataFrame(self.colValuesDict)
        df.to_excel("output.xlsx", index=False, columns=moduleColumns)

    # 通过某个值获取DataFrame中该值对应的行列数组
    def GetIndexes(self, dfObj, value):
        listOfPos = list()
        # 得到值均为布尔值的DataFrame
        # result = dfObj.applymap(lambda x: (value in str(x)))
        result = dfObj.applymap(lambda x: self.MatchRemoveSpace(x, value))

        # 得到包含该值的列的数组
        seriesObj = result.any()

        columnNames = list(seriesObj[seriesObj == True].index)
        # 得到包含该值的行的数组
        for col in columnNames:
            rows = list(result[col][result[col] == True].index)
            colIndex = result.columns.get_loc(col)
            for row in rows:
                listOfPos.append((row, colIndex, col))
        return listOfPos


    def ReadExcel(self, fPath):
        if(len(fPath) > 0):
            # 忽略前两行
            data = pd.read_excel(io=fPath, sheet_name=None, header=0, skiprows=2)
            self.resData = data['明细']