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

dbcBADEFContext = "BA_DEF_  \"BusType\" STRING ;\nBA_DEF_ BO_ \"GenMsgCycleTime\" INT 0 65535;\n"
dbcBADEFDEFContext = "BA_DEF_DEF_  \"BusType\" \"CAN\";\n"

colMsgName = 0
colID = 2
colMsgSendType = 3
colMsgSendCycle = 4
colMsgDlc = 5
colSgName = 6
colSgDescription = 7
colSgByteOrder = 8
colSgStartBit = 10
colSgBitLen = 12
colSgDataType = 13
colSgFactor = 14
colSgOffset = 15
colSgMin = 16
colSgMax = 17
colSgUnit = 23
colNodeStart = 28

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

        nodes = self.matrixTable.row_values(0)[colNodeStart:]
        noRow = 1
        while noRow < self.matrixTable.nrows:
            noRowData = self.matrixTable.row_values(noRow)
            noRow+=1

            if noRowData[0] != "":
                numID = eval(noRowData[colID])
                if numID > 0x7ff:
                    numID = numID | (1 << 31)
                canID = str(numID)
                MsgName = noRowData[colMsgName]+":"
                MsgDlc = str(int(noRowData[colMsgDlc]))
                nodeIndex = 0
                for iter in noRowData[colNodeStart:]:
                    if(iter == 'S'):
                        node = nodes[nodeIndex] + "\n"
                        break
                    nodeIndex += 1
                BOContext = " ".join(["\nBO_", canID, MsgName, MsgDlc, node])
                self.dbc_fd.write(BOContext)
            else:
                # 信号属性获取
                SignalName = re.sub("\W+","",noRowData[colSgName])
                SignalStartBit = int(noRowData[colSgStartBit])
                SignalBitLenth = int(noRowData[colSgBitLen])
                SignalOffset = 0 if isEmpty(noRowData[colSgOffset]) else getVal(noRowData[colSgOffset])
                SignalByteOrder = "1" if noRowData[colSgByteOrder] == "Intel" else "0"
                SignalFactor = 1 if isEmpty(noRowData[colSgFactor]) else getVal(noRowData[colSgFactor])
                if noRowData[colSgDataType] == "signed":
                    SignalDataType = "-"
                    SignalMin = ((-pow(2,SignalBitLenth-1))*SignalFactor + SignalOffset) if isEmpty(noRowData[colSgMin]) else getVal(noRowData[colSgMin])
                    SignalMax = ((pow(2, SignalBitLenth-1)-1)*SignalFactor + SignalOffset) if isEmpty(noRowData[colSgMax]) else getVal(noRowData[colSgMax])
                else:
                    SignalDataType = "+"
                    SignalMin = SignalOffset if isEmpty(noRowData[colSgMin]) else getVal(noRowData[colSgMin])
                    SignalMax = ((pow(2, SignalBitLenth)-1)*SignalFactor + SignalOffset) if isEmpty(noRowData[colSgMax]) else getVal(noRowData[colSgMax])
                SignalUnit = "\"" + getUnit(noRowData[colSgUnit]) + "\""
                # 预处理
                bitContext = str(SignalStartBit) + "|" + str(SignalBitLenth) + "@" + SignalByteOrder + SignalDataType
                # valueContext = "(" + str(SignalFactor) + "," + str(SignalOffset) + ")"
                # limitContext = "[" + str(SignalMin) + "|" + str(SignalMax) + "]"
                valueContext = "({:g},{:g})".format(SignalFactor, SignalOffset)
                limitContext = "[{:g}|{:g}]".format(SignalMin, SignalMax)
                nodeIndex = 0
                for iter in noRowData[colNodeStart:]:
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
            if noRowData[colMsgName] != "":
                numID = eval(noRowData[colID])
                if numID > 0x7ff:
                    numID = numID | (1 << 31)
                canID = str(numID)
            elif noRowData[colSgDescription] != "":
                SignalName = noRowData[colSgName]
                SignalDescribe = "\"" + noRowData[colSgDescription]+ "\";\n"
                newContext = " ".join(["CM_", "SG_", canID, SignalName, SignalDescribe])
                self.dbc_fd.write(newContext)
            noRow+=1

        self.dbc_fd.write(dbcBADEFContext)
        self.dbc_fd.write(dbcBADEFDEFContext)

        noRow = 1
        while noRow < self.matrixTable.nrows:
            noRowData = self.matrixTable.row_values(noRow)
            if noRowData[colMsgName] != "":
                numID = eval(noRowData[colID])
                if numID > 0x7ff:
                    numID = numID | (1 << 31)
                canID = str(numID)
                if noRowData[colMsgSendType] == "Periodic":
                    SignalCycle = str(int(noRowData[colMsgSendCycle]))
                    newContext = " ".join(["BA_", "\"GenMsgCycleTime\"", "BO_", canID, SignalCycle, ";\n"])
                    self.dbc_fd.write(newContext)
            noRow+=1

if __name__ == "__main__":
    excel = ExcelLoad("Template.xls")
    excel.Convert()