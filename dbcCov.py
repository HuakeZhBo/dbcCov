import dbc2excel
import excel2dbc
import sys
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Arguments Error!")
        exit()
    convType = sys.argv[1]
    fileIn = sys.argv[2]
    if convType == "dbc" :
        dbc = dbc2excel.DbcLoad(fileIn)
        dbc.Convert()
    elif convType == "excel":
        excel = excel2dbc.ExcelLoad(fileIn)
        excel.Convert()