import DataHandler as dh
import os

eh = dh.ExcelHandler()
# 要用path.join,因为跨平台路径不同
fileFolderPath = os.path.join(os.getcwd(), "ResFiles")
# 也可以使用相对目录
# fileFolderPath = "ResFiles"
eh.ReadFileFolder(fileFolderPath)
eh.ReadTestModuleExcel("模板.xlsx")

eh.SearchContentFromFiles()
eh.BuildDataFrame()