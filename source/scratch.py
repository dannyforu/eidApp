import eidApp
import os
import time
import random
from eidApp import Name2Code


filenum = 0
for root, dirs, files in os.walk(r"D:\data\import"):
    print("enter the dir " + root)
    if "学生" in root:
        # 当前路径下所有非目录子文件
        for file in files:
            (shotname, extension) = os.path.splitext(file)
            if extension == ".csv":
                print("处理文件：" + file)
                eidapp1 = eidApp.eidAppFile_student()
                eidapp1.loadFromFile(root + "\\" + file)
                eidapp1.dealdata()
                filenum += 1
            elif extension == ".xls" or extension == ".xlsx":
                print("暂不处理处理文件：" + file)
            else:
                print("不支持" + file)
print("共处理" + str(filenum) + "个文件")

