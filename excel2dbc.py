import xlrd
import re

dbcHeaderContext="VERSION \"\"\n\n\nNS_ :\n\tNS_DESC_\n\
	CM_\n\
	BA_DEF_\n\
	BA_\n\
	VAL_\n\
	CAT_DEF_\n\
	CAT_\n\
	FILTER\n\
	BA_DEF_DEF_\n\
	EV_DATA_\n\
	ENVVAR_DATA_\n\
	SGTYPE_\n\
	SGTYPE_VAL_\n\
	BA_DEF_SGTYPE_\n\
	BA_SGTYPE_\n\
	SIG_TYPE_REF_\n\
	VAL_TABLE_\n\
	SIG_GROUP_\n\
	SIG_VALTYPE_\n\
	SIGTYPE_VALTYPE_\n\
	BO_TX_BU_\n\
	BA_DEF_REL_\n\
	BA_REL_\n\
	BA_DEF_DEF_REL_\n\
	BU_SG_REL_\n\
	BU_EV_REL_\n\
	BU_BO_REL_\n\
	SG_MUL_VAL_\n\n\
BS_:\n\n"

dbcEOFContext = "BA_DEF_  \"BusType\" STRING ;\n\
BA_DEF_DEF_  \"BusType\" \"CAN\";\n\n"

rowMsgName = 0
rowID = 2
rowMsgDlc = 5
rowSgName = 6
rowSgDescription = 7
rowSgByteOrder = 8
rowSgStartBit = 10
rowSgBitLen = 12
rowSgDataType = 13
rowSgFactor = 14
rowSgOffset = 15
rowSgMin = 16
rowSgMax = 17
rowSgUnit = 23
rowNodeStart = 28

def isEmpty(x):
    if x == '/' or x == '':
        return 1
    else:
        return 0
    
def getVal(x):
    if isinstance(x,str):
        return eval(x)
    else:
        return x

def getUnit(x):
    if x == '/' or x == '-':
        return ''
    else:
        return x

class ExcelLoad(object):
    def __init__(self, filepath):
        self.excel_path = filepath
        self.excle_name = filepath.split("\\")[-1]
        self.dbc_name = self.excle_name.split(".")[0] + ".dbc"

    def Convert(self):
        # 打开文件
        self.excel_fd = xlrd.open_workbook(self.excel_path)
        self.dbc_fd = open(self.dbc_name , "w+")
        # 获取工作表
        self.matrixTable = self.excel_fd.sheet_by_name('Matrix')
        self.nodeTable = self.excel_fd.sheet_by_name('Nodes')

        self.dbc_fd.write(dbcHeaderContext)

        # 写入结点列表
        nodes = self.nodeTable.col_values(0, 1)
        nodeContext = "BU_:"
        for node in nodes:
            nodeContext = " ".join([nodeContext, node])
        nodeContext = " ".join([nodeContext, '\n\n'])
        self.dbc_fd.write(nodeContext)

        nodes = self.matrixTable.row_values(0)[rowNodeStart:]
        noRow = 1
        while noRow < self.matrixTable.nrows:
            noRowData = self.matrixTable.row_values(noRow)
            noRow+=1

            if noRowData[0] != "":
                intID = str(eval(noRowData[rowID]))
                MsgName = noRowData[rowMsgName]+":"
                MsgDlc = str(int(noRowData[rowMsgDlc]))
                nodeIndex = 0
                for iter in noRowData[rowNodeStart:]:
                    if(iter == 'S'):
                        node = nodes[nodeIndex] + "\n"
                        break
                    nodeIndex += 1
                BOContext = " ".join(["\nBO_", intID, MsgName, MsgDlc, node])
                self.dbc_fd.write(BOContext)
            else:
                # 信号属性获取
                SignalName = re.sub("\W+","",noRowData[rowSgName])
                SignalStartBit = int(noRowData[rowSgStartBit])
                SignalBitLenth = int(noRowData[rowSgBitLen])
                SignalOffset = 0 if isEmpty(noRowData[rowSgOffset]) else getVal(noRowData[rowSgOffset])
                SignalByteOrder = "1" if noRowData[rowSgByteOrder] == "Intel" else "0"
                SignalFactor = 1 if isEmpty(noRowData[rowSgFactor]) else getVal(noRowData[rowSgFactor])
                if noRowData[rowSgDataType] == "signed":
                    SignalDataType = "-"
                    SignalMin = ((-pow(2,SignalBitLenth-1))*SignalFactor + SignalOffset) if isEmpty(noRowData[rowSgMin]) else getVal(noRowData[rowSgMin])
                    SignalMax = ((pow(2, SignalBitLenth-1)-1)*SignalFactor + SignalOffset) if isEmpty(noRowData[rowSgMin]) else getVal(noRowData[rowSgMax])
                else:
                    SignalDataType = "+"
                    SignalMin = SignalOffset if isEmpty(noRowData[rowSgMin]) else getVal(noRowData[rowSgMin])
                    SignalMax = ((pow(2, SignalBitLenth)-1)*SignalFactor + SignalOffset) if isEmpty(noRowData[rowSgMax]) else getVal(noRowData[rowSgMax])
                SignalUnit = "\"" + getUnit(noRowData[rowSgUnit]) + "\""
                # 预处理
                bitContext = str(SignalStartBit) + "|" + str(SignalBitLenth) + "@" + SignalByteOrder + SignalDataType
                # valueContext = "(" + str(SignalFactor) + "," + str(SignalOffset) + ")"
                # limitContext = "[" + str(SignalMin) + "|" + str(SignalMax) + "]"
                valueContext = "({:g},{:g})".format(SignalFactor, SignalOffset)
                limitContext = "[{:g}|{:g}]".format(SignalMin, SignalMax)
                nodeIndex = 0
                for iter in noRowData[rowNodeStart:]:
                    if(iter == 'R'):
                        node = nodes[nodeIndex] + "\n"
                        break
                    nodeIndex += 1
                # 拼接写入
                SGContext = " ".join([" SG_", SignalName, ":", bitContext, valueContext, limitContext, SignalUnit, node])
                self.dbc_fd.write(SGContext)

        # 写入节点描述
        noRow = 1
        while noRow < self.nodeTable.nrows:
            noRowData = self.nodeTable.row_values(noRow)
            if(noRowData[1] != ""):
                nodeDesc = "\"" + noRowData[1] + "\";\n"
                nodeDescrbeContext = " ".join(['CM_', 'BU_', noRowData[0], nodeDesc])
                self.dbc_fd.write(nodeDescrbeContext)
            noRow += 1

        # 写入信号描述
        noRow = 1
        while noRow < self.matrixTable.nrows:
            noRowData = self.matrixTable.row_values(noRow)
            if noRowData[rowMsgName] != "":
                intID = str(eval(noRowData[rowID]))
            elif noRowData[rowSgDescription] != "":
                SignalName = noRowData[rowSgName]
                SignalDescribe = "\"" + noRowData[rowSgDescription]+ "\";\n"
                newContext = " ".join(["CM_", "SG_", intID, SignalName, SignalDescribe])
                self.dbc_fd.write(newContext)
            noRow+=1

        self.dbc_fd.write(dbcEOFContext)

if __name__ == "__main__":
    excel = ExcelLoad("Template.xls")
    excel.Convert()