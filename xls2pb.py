from xml.dom.minidom import parse
import xlrd # for read excel
import sys
import os

class LogHelp :
    """日志辅助类"""
    _logger = None
    _close_imme = True

    @staticmethod
    def set_close_flag(flag):
        LogHelp._close_imme = flag

    @staticmethod
    def _initlog():
        import logging

        LogHelp._logger = logging.getLogger()
        logfile = 'convert.log'
        hdlr = logging.FileHandler(logfile)
        formatter = logging.Formatter('%(asctime)s|%(levelname)s|%(lineno)d|%(funcName)s|%(message)s')
        hdlr.setFormatter(formatter)
        LogHelp._logger.addHandler(hdlr)
        LogHelp._logger.setLevel(logging.NOTSET)
        # LogHelp._logger.setLevel(logging.WARNING)

        LogHelp._logger.info("logger is inited!")

    @staticmethod
    def get_logger() :
        if LogHelp._logger is None :
            LogHelp._initlog()

        return LogHelp._logger

    @staticmethod
    def close() :
        if LogHelp._close_imme:
            import logging
            if LogHelp._logger is None :
                return
            logging.shutdown()

# log macro
LOG_DEBUG=LogHelp.get_logger().debug
LOG_INFO=LogHelp.get_logger().info
LOG_WARN=LogHelp.get_logger().warn
LOG_ERROR=LogHelp.get_logger().error

TAP_BLANK_NUM = "    "
FIELD_SINGULAR_TYPE = "singular"
FIELD_REPEATED_TYPE = "repeated"

pbFileList = []

class MsgField:
    def __init__(self, fieldName, fieldRule, fieldType, fieldCName, fieldDesc):
        self.fieldName = fieldName
        self.fieldRule = fieldRule
        self.fieldType = fieldType
        self.fieldCName = fieldCName
        self.fieldDesc = fieldDesc

class MsgStruct:
    def __init__(self, msgName, xlsName):
        self.msgName = msgName
        self.xlsName = xlsName
        self.fieldMap = {}
    def GetFieldDefine(self, fieldCName):
        if fieldCName in self.fieldMap:
            return self.fieldMap[fieldCName]

class PbFile:
    def __init__(self, fileName):
        self.fileName = fileName
        self.msgMap = {}


class XmlParser:
    def __init__(self, xml_file_path):
        self.dom = parse(xml_file_path)
    
    def Parse(self):
        data = self.dom.documentElement
        fileName = data.getAttribute('name')
        pbFileContext = PbFile(fileName)
        idx = 0
        messageList = data.getElementsByTagName('message')
        for message in messageList:
            # 获取标签属性值
            msgName = message.getAttribute('name')
            xlsName = message.getAttribute('cname')
            msgStruct = MsgStruct(msgName, xlsName)
            fieldList = message.getElementsByTagName('field')
            for field in fieldList:
                idx += 1
                fieldName = field.getAttribute('name')
                fieldRule = field.getAttribute('rule')
                fieldType = field.getAttribute('type')
                fieldCName = field.getAttribute('cname')
                fieldDesc = field.getAttribute('desc')
                msgField = MsgField(fieldName, fieldRule, fieldType, fieldCName, fieldDesc)
                msgStruct.fieldMap[fieldCName] = msgField
            pbFileContext.msgMap[xlsName] = msgStruct
        pbFileList.append(pbFileContext)

class PbFileParser:
    def __init__(self):
        # 读取模板文件
        file = open("template.proto", mode='r', encoding='utf8')
        self.pbFileContext = file.read()
        file.close()

    def Parse(self):
        for pbFile in pbFileList:
            context = self.pbFileContext
            for msg in pbFile.msgMap.values():
                # 获取标签属性值
                idx = 0
                message = ""
                for field in msg.fieldMap.values():
                    idx += 1
                    if field.fieldDesc.count("\n") > 1:
                        if field.fieldDesc[-1] != '\n':
                            field.fieldDesc = field.fieldDesc + "\n"
                        field.fieldDesc = strings.Replace(field.fieldDesc, "\n", "\n"+TAP_BLANK_NUM, -1)
                        field.fieldDesc = strings.Replace(field.fieldDesc, "\n\n", "\n", -1)
                        message += TAP_BLANK_NUM + "/** " + field.fieldDesc + TAP_BLANK_NUM + "*/\n"
                    else:
                        message += TAP_BLANK_NUM + "/** " + field.fieldDesc + " */\n"

                    if field.fieldRule == FIELD_SINGULAR_TYPE:
                        message += TAP_BLANK_NUM + field.fieldType + " " + field.fieldName + " = " + str(idx) + ";\n"
                    else:
                        message += TAP_BLANK_NUM + field.fieldRule + " " + field.fieldType + " " + field.fieldName + " = " + str(idx) + ";\n"
                message = "message " + msg.msgName + "{\n" + message + "}\n"
                context += message
                if msg.xlsName != "":
                    context += "\nmessage " + msg.msgName + "List {\n    repeated " + msg.msgName + " data = 1;\n}\n\n"

            self.WritePbFile(pbFile.fileName, context)
            self.GenPbFile(pbFile.fileName)

    def WritePbFile(self, fileName, fileContext) :
        fileName = fileName + ".proto"
        file = open(fileName, 'w', encoding='utf-8')
        file.write(fileContext)
        file.close()
    
    def GenPbFile(self, fileName) :
        fileName = fileName + ".proto"
        try :
            command = "protoc --python_out=./ " + fileName
            os.system(command)
        except BaseException :
            print("protoc failed!")
            raise

class XlsParser:
    def __init__(self, xlsx_file_path):
        self._xls_file_path = xlsx_file_path

        try :
            self._workbook = xlrd.open_workbook(self._xls_file_path)
        except BaseException:
            LOG_DEBUG("open xls file(%s) failed!", self._xls_file_path)
            raise

    def Parse(self) :
        """对外的接口:解析数据"""
        
        for sheet in self._workbook.sheets():
            if "Sheet" in sheet.name:
               LOG_DEBUG("Sheet(%s) is not used", sheet.name)
               continue
            try:
                self._sheet = self._workbook.sheet_by_name(sheet.name)
            except BaseException:
                LOG_ERROR("open sheet(%s) failed!", sheet.name)
                raise

            # 获取sheet对应的pb文件       
            pbFile = GetPbFile(sheet.name)
            # 获取sheet对应的pb msg定义
            msgDefine = pbFile.msgMap[sheet.name]
            try:
                self._module_name = pbFile.fileName + "_pb2"
                sys.path.append(os.getcwd())
                exec('from ' + self._module_name + ' import *')
                self._module = sys.modules[self._module_name]
            except BaseException:
                LOG_ERROR("load module(%s) failed", self._module_name)
                raise

            # 获取sheet对应的pb类型
            try:
                dataList = getattr(self._module, msgDefine.msgName + 'List')()      
            except BaseException:
                LOG_ERROR("%s getattr %s failed", self._module, msgDefine)
                raise

            self._row_count = len(self._sheet.col_values(0))
            self._col_count = len(self._sheet.row_values(0))
            self._row = 0
            self._col = 0
            # 逐行解析数据到pb类型
            for self._row in range(1, self._row_count):
                data = dataList.data.add()
                for self._col in range(0, self._col_count):
                    # 解析列对应的pb字段, 首行对应对应xml当中field中的cname属性
                    fieldCName = str(self._sheet.cell_value(0, self._col))
                    fieldCNameList = fieldCName.split("_")
                    ParseField(msgDefine, fieldCNameList, data, self._sheet.cell_value(self._row, self._col))

            LOG_DEBUG("%s config data %s", pbFile.fileName, str(dataList))
            self._WriteReadableData2File(pbFile.fileName, str(dataList))
            data = dataList.SerializeToString()
            self._WriteData2File(pbFile.fileName, data)

    # 将序列化过的pb数据写入文件
    def _WriteData2File(self, fileName, data) :
        file_name = fileName + ".bin"#self._proto_name.lower() + ".bin"
        file = open(file_name, 'wb+')
        file.write(data)
        file.close()

    # 将可识别的pb数据写入文件
    def _WriteReadableData2File(self, fileName, data) :
        file_name = fileName + ".txt"#self._proto_name.lower() + ".txt"
        file = open(file_name, 'w', encoding='utf-8')
        file.write(data)
        file.close()
def GetPbFile(sheetName):
    for pbFile in pbFileList:
        if sheetName in pbFile.msgMap:
            return pbFile

def GetMsgDefine(msgName):
    for pbFile in pbFileList:
        for msg in pbFile.msgMap.values():
            if msg.msgName == msgName:
                return msg

pbBaseTypeList = ["int32", "int64", "uint32", "uint64", "sint32", "sint64", "fixed32", "fixed64", "sfixed32", "sfixed64", "double", "float", "string", "bytes"]
# 解析Field
def ParseField(msgDefine, fieldCNameList, data, cellValue):
    fieldDefine = msgDefine.GetFieldDefine(fieldCNameList[0])
    # 判断是否是proto的基础类型
    bPbBaseType = IsPbBaseType(fieldDefine.fieldType)
    if not bPbBaseType:
        if fieldDefine.fieldRule == FIELD_REPEATED_TYPE:
            if len(fieldCNameList) > 1:
                fieldData = data.__getattribute__(fieldDefine.fieldName).add()
                ParseField(GetMsgDefine(fieldDefine.fieldType), fieldCNameList[1:], fieldData, cellValue)
            else :
                fieldData.append(GetFieldValue(fieldDefine.fieldType, cellValue))
        else:
            if len(fieldCNameList) > 1:
                fieldData = data.__getattribute__(fieldDefine.fieldNam)
                ParseField(GetMsgDefine(fieldDefine.fieldType), fieldCNameList[1:], fieldData, cellValue)
            else:
                data.__setattr__(fieldDefine.fieldName, GetFieldValue(fieldDefine.fieldType,cellValue))
    else:
        if fieldDefine.fieldRule == FIELD_REPEATED_TYPE:
            data.__getattribute__(fieldDefine.fieldName).append(GetFieldValue(fieldDefine.fieldType,cellValue))
        else:
            data.__setattr__(fieldDefine.fieldName, GetFieldValue(fieldDefine.fieldType,cellValue))

# 值转换
def GetFieldValue(field_type, field_value) :
    """将pb类型转换为python类型"""
    try:
        if field_type == "int32" or field_type == "int64"\
                or  field_type == "uint32" or field_type == "uint64"\
                or field_type == "sint32" or field_type == "sint64"\
                or field_type == "fixed32" or field_type == "fixed64"\
                or field_type == "sfixed32" or field_type == "sfixed64" :
                    if len(str(field_value).strip()) <=0 :
                        return None
                    else :
                        return int(field_value)
        elif field_type == "double" or field_type == "float" :
                if len(str(field_value).strip()) <=0 :
                    return None
                else :
                    return float(field_value)
        elif field_type == "string" :
            field_value = str(field_value)
            if len(field_value) <= 0 :
                return None
            else :
                return field_value
        # elif field_type == "bytes" :
        #     # field_value = unicode(field_value).encode('utf-8')
        #     # if len(field_value) <= 0 :
        #     #     return None
        #     # else :
        #     #     return field_value
        else :
            return None
    except BaseException:
        LOG_ERROR("parse cell(%u, %u) error, please check it, maybe type is wrong.", row, col)
        raise

# 判断是否是pb基础类型
def IsPbBaseType(fieldType):
    for baseType in pbBaseTypeList:
        if fieldType == baseType:
            return 1
    return 0

files= os.listdir("./")

for file in files: #遍历文件夹
     if not os.path.isdir(file) and file.count(".xml"):
        LOG_DEBUG("parse %s file", file)
        xmlParser = XmlParser(file)
        xmlParser.Parse()

pbParser = PbFileParser()
pbParser.Parse()

for file in files: #遍历文件夹
     if not os.path.isdir(file) and file.count(".xls"):
        LOG_DEBUG("parse %s file", file)
        xlsParser = XlsParser(file)
        xlsParser.Parse()
