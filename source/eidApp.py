import csv
import pandas
import os, shutil
import numpy as np
import time

class rec_eidAppFile_student:
    def loaddata(self, line):
        self.name = line[1]
        self.studyid = line[2]
        self.ctid = line[3]
        self.sex = line[4]
        return

class Name2Code:
    def __init__(self, pFileName, pNameFlag=0):
        self._filename = pFileName
        self._defaultCode = ""
        if pNameFlag == 0:
            NameCol = 1
            CodeCol = 0
        else:
            NameCol = 0
            CodeCol = 1

        self._codetable = {}
        file = open(self._filename, "r", encoding="gbk")
        while 1:
            rowdata = file.readline().strip("\r").strip("\n")
            if rowdata == "":
                break
            data = rowdata.split()
            self._codetable[data[NameCol]] = data[CodeCol]

    def getCode(self, pName):
        if pName in self._codetable:
            return self._codetable[pName]
        else:
            return self._defaultCode

    def setDefaultValue(self, pCode):
        self._defaultCode=pCode

class eidAppFile_student:
    def __init__(self):
        return

    def loadFromFile(self, pfile):
        self._file = pfile
        (self._filepath, self._filename) = os.path.split(pfile)
        self._csvfile = open(self._file, "r")
        self._lines = csv.DictReader(self._csvfile)
        return

    def search(self, path, name):
        for root, dirs, files in os.walk(path):  # path 为根目录
            if name in files:
                flag = 1  # 判断是否找到文件
                root = str(root)
                return os.path.join(root+"\\")
        return "-1"

    def printdata(self):
        for recdata in self._lines:
            print(recdata[0])
        return

    def dealdata(self):
        today = time.strftime("%Y%m%d", time.localtime(time.time()))
        headers = ["姓名", "学籍号", "性别", "出生日期", "民族", "国籍", "政治面貌", "身份证件类型", "身份证件号", "有效起始日期", "有效截止日期", "学校名称",
                   "学校标识码", "籍贯", "入学年月", "学生来源", "学生类别", "年级", "班号", "港澳台侨外", "专业名称", "健康状况", "出生地", "家庭地址", "邮政编码",
                   "联系电话", "电子邮件地址", "户口性质", "入学方式", "就读方式", "是否进城务工人员子女", "是否孤儿", "是否烈士优抚子女", "是否需要申请资助", "是否享受一补",
                   "随班就读", "是否残疾", "联招合作类型", "联招合作学校机构代码", "监护人姓名", "关系", "身份证件类型", "身份证件号码", "有效起始日期", "有效截止日期",
                   "监护人工作单位", "监护人通讯地址", "邮政编码", "联系电话"]
        coltype = {
            "姓名": np.object, "学籍号": np.object, "性别": np.object, "出生日期": np.object, "民族": np.object, "国籍": np.object,
            "政治面貌": np.object, "身份证件类型": np.object, "身份证件号": np.object, "有效起始日期": np.object, "有效截止日期": np.object,
            "学校名称": np.object,
            "学校标识码": np.object, "籍贯": np.object, "入学年月": np.object, "学生来源": np.object, "学生类别": np.object,
            "年级": np.object, "班号": np.object, "港澳台侨外": np.object, "专业名称": np.object, "健康状况": np.object,
            "出生地": np.object, "家庭地址": np.object, "邮政编码": np.object,
            "联系电话": np.object, "电子邮件地址": np.object, "户口性质": np.object, "入学方式": np.object, "就读方式": np.object,
            "是否进城务工人员子女": np.object, "是否孤儿": np.object, "是否烈士优抚子女": np.object, "是否需要申请资助": np.object,
            "是否享受一补": np.object,
            "随班就读": np.object, "是否残疾": np.object, "联招合作类型": np.object, "联招合作学校机构代码": np.object, "监护人姓名": np.object,
            "关系": np.object, "身份证件类型": np.object, "身份证件号码": np.object, "有效起始日期": np.object, "有效截止日期": np.object,
            "监护人工作单位": np.object, "监护人通讯地址": np.object, "邮政编码": np.object, "联系电话": np.object}
        # init the dictionary table
        mzmap = Name2Code("D:\\data\\code\\民族.txt")
        mzmap.setDefaultValue("01")
        gjmap = Name2Code("D:\\data\\code\\国家代码.txt")
        gjmap.setDefaultValue("156")
        zzmm = Name2Code("D:\\data\\code\\政治面貌.txt")
        zzmm.setDefaultValue("13")
        xsly = Name2Code("D:\\data\\code\\学生来源.txt")
        xsly.setDefaultValue("1")
        gatw = Name2Code("D:\\data\\code\\港澳台侨外.txt")
        gatw.setDefaultValue("00")

        wfile = "D:\data\output\\" + self._filename
        if os.path.exists(wfile):
            (shortname, extension) = os.path.splitext(self._filename)
            wfile = "D:\data\output\\" + shortname + str(int(time.time()*1000)) + ".csv"
        f = open(wfile, 'w', newline='')
        # 标头在这里传入，作为第一行数据
        writer = csv.DictWriter(f, headers)
        writer.writeheader()
        studata3 = {}
        # 写入headers, 兼容申请表格式中第二行为实际HEADER的要求
        for i in headers:
            studata3[i] = i
        writer.writerow(studata3)
        # 写入headers完成
        studata2 = {}
        for studata1 in self._lines:
            # 基础信息处理
            studata2["姓名"] = studata1["学生姓名"].replace("\t", "").replace(" ", "")
            studata2["学籍号"] = studata1["学籍号"].replace("\t", "").replace(" ", "")
            if studata1["性别"] == "男":
                studata2["性别"] = "1"
            elif studata1["性别"] == "女":
                studata2["性别"] = "2"
            else:
                studata2["性别"] = "9"
            studata2["出生日期"] = studata1["出生日期"].replace("\t", "").replace(" ", "")
            # 转换民族编码
            studata2["民族"] = mzmap.getCode(studata1["民族"].replace("\t", "").replace(" ", ""))
            # 转换国籍编码
            studata2["国籍"] = gjmap.getCode(studata1["国籍地区"].replace("\t", "").replace(" ", ""))
            # 转换政治面貌编码
            studata2["政治面貌"] = zzmm.getCode(studata1["政治面貌"].replace("\t", "").replace(" ", ""))
            '''
            if "群众" in studata1["政治面貌"]:
                studata2["政治面貌"] = "13"
            elif "共产党党员" in studata1["政治面貌"]:
                studata2["政治面貌"] = "01"
            else:
                studata2["政治面貌"] = "13"
            '''
            studata2["身份证件类型"] = "1"
            studata2["身份证件号"] = studata1["身份证件号"].replace("\t", "").replace(" ", "")

            '''
            处理照片
            '''
            picfileName = studata2["身份证件号"]+".jpg"
            destPicDir = "D:\data\output\DP\\"
            picfiledir= "D:\data\pic\DP\\"
            sourcePicDir = self.search(picfiledir, picfileName)
            if sourcePicDir != '-1':
                sourcePic = sourcePicDir + picfileName
                destPic = destPicDir + "DP-"+studata2["学籍号"]+"-"+today+".jpg"
                shutil.copy(sourcePic, destPic)
            else:
                continue  #若没有照片，不处理其他数据

            # 未获取起始日期、终止日期数据，设为缺省值
            studata2["有效起始日期"] = "20180101"
            studata2["有效截止日期"] = "20180101"
            studata2["学校名称"] = studata1["学校名称"].replace("\t", "").replace(" ", "")
            studata2["学校标识码"] = studata1["学校标识码"].replace("\t", "").replace(" ", "")
            # 根据身份证号确定籍贯
            studata2["籍贯"] = studata2["身份证件号"][0:6]
            studata2["入学年月"] = studata1["入学年月"].replace("\t", "").replace(" ", "")
            # 转换学生来源编码
            studata2["学生来源"] = xsly.getCode(studata1["学生来源"].replace("\t", "").replace(" ", ""))
            '''
            studata2["学生来源"] = studata1["学生来源"].replace("\t", "").replace(" ", "")
            if studata2["学生来源"] == "正常入学":
                studata2["学生来源"] = "1"
            elif studata2["学生来源"] == "借读":
                studata2["学生来源"] = "2"
            else:
                studata2["学生来源"] = "1"
            '''
            studata2["年级"] = studata1["年级"].replace("\t", "").replace(" ", "")
            studata2["班号"] = studata1["班级"].replace("\t", "").replace(" ", "")
            # 没有学生类别数据，根据年级判断
            if "初中" in studata1["年级"]:
                studata2["学生类别"] = "31100"
            elif "小学" in studata1["年级"]:
                studata2["学生类别"] = "21100"
            else:
                studata2["学生类别"] = "0"
            # 转换港澳台侨外编码
            if "否" in studata1["港澳台侨外"]:
                studata2["港澳台侨外"] = "00"
            else:
                studata2["港澳台侨外"] = gatw.getCode(studata1["港澳台侨外"].replace("\t", "").replace(" ", ""))
            '''    
            elif "香港" in studata1["港澳台侨外"]:
                studata2["港澳台侨外"] = "01"
            elif "澳门" in studata1["港澳台侨外"]:
                studata2["港澳台侨外"] = "03"
            elif "台湾" in studata1["港澳台侨外"]:
                studata2["港澳台侨外"] = "05"
            else:
                studata2["港澳台侨外"] = "99"
            '''
            # 基本信息处理完毕
            writer.writerow(studata2)


        # 转换成XLS文件
        f.close()
        csvfile = pandas.read_csv(wfile, encoding='gbk', dtype=coltype)
        (wpath, wfilename) = os.path.split(wfile)
        (shortname, extension) = os.path.splitext(wfilename)
        xlsfile = "D:\data\\output\\" + shortname + ".xlsx"
        csvfile.to_excel(xlsfile, sheet_name='sheet1', index=False)

        return