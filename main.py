# 删除表格中的重复行（根据列判断）
from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell
from odf.text import P
import pandas as pd
import os

# 配置
class config():
    # 文件路径
    fileInputPath = "删除表格的重复行/OpenDocument Spreadsheet.csv"
    # 输出文件路径
    fileOutputPath = "删除表格的重复行/OpenDocument Spreadshee1t.csv"
    # 需要判断的列的名 二选一
    rowName = ["测试"]
    # 需要判断的列的序号 二选一 优先
    rowNo = []

# 删除重复行
def deleteSame(df):
    if config.rowNo:
        # 保留重复行的第一行
        duplicated_rows = df[df.duplicated(subset=df.columns[config.rowNo], keep='first')]
    else:
        duplicated_rows = df[df.duplicated(subset=config.rowName, keep='first')]

    # 删除包含相同值的行
    df_cleaned = df.drop(duplicated_rows.index)

    # 返回删除重复行后的 DataFrame
    return df_cleaned

# 写入ods文件
def odfWrite(df, fileOutputPath):
    doc = OpenDocumentSpreadsheet()

    # 创建一个新表格
    table = Table(name="Sheet1")

    # 获取表头
    tr = TableRow()
    for col in df.columns:
        tc = TableCell()
        tc.addElement(P(text=col))
        tr.addElement(tc)
    table.addElement(tr)

    # 从df中获取每一行
    for row in df.itertuples(index=False):
        tr = TableRow()
        for cell in row:
            tc = TableCell()
            tc.addElement(P(text=str(cell)))
            tr.addElement(tc)
        table.addElement(tr)

    # 将表格加入文件中
    doc.spreadsheet.addElement(table)

    # 保存ods文件
    doc.save(fileOutputPath)

# 读取表格文件
def readFile(fileInputPath, fileType):
    if fileType in ('xls', 'xlsx'):
        df = pd.read_excel(fileInputPath)
    elif fileType == 'ods':
        df = pd.read_excel(fileInputPath,engine='odf')
    elif fileType == 'csv':
        df = pd.read_csv(fileInputPath)
    return df

# 写入表格文件
def writeFile(df, fileOutputPath, fileType):
    try:
        if fileType in ('xls', 'xlsx'):
            if fileType == 'xls':
                # 新版本的pandas不再支持xls写入
                df.to_excel(fileOutputPath+'.xlsx', sheet_name="Sheet1")
            else:
                df.to_excel(fileOutputPath, sheet_name="Sheet1", engine='xlwt')
        elif fileType == 'ods':
            odfWrite(df, fileOutputPath)
        elif fileType == 'csv':
            df.to_csv(fileOutputPath)
        else:
            print("Unexplan name")
            return False
        return True
    except:
        return False

if __name__ == '__main__':
    fileType = os.path.splitext(config.fileInputPath)[1].replace('.','')

    df = readFile(config.fileInputPath, fileType)
    df_cleaned = deleteSame(df)

    print(df)
    print("Result Of Cleaned")
    print(df_cleaned)

    print(f"Result Of Write is {writeFile(df_cleaned, config.fileOutputPath, fileType)}")
