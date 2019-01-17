#!/usr/bin/env python
# _*_ UTF-8 _*_
'''
@project:copy-excel
@author:xiaofei
'''
import os
import os.path
import shutil
import time,  datetime

def copyFileto(sourcefile, title_list):
    # 创建相对路径的文件夹
    curPath = os.getcwd()
    tempPath = 'file_target'
    targetPath = curPath + os.path.sep + tempPath
    if not os.path.exists(targetPath):
        os.makedirs(targetPath)
    # 复制文件
    for i in range(len(title_list)):
        target = targetPath + "/" + title_list[i] + ".xlsx"
        shutil.copy(sourcefile, target)

if __name__ == '__main__':
    sourcefile = 'C:/Users/XXX/Desktop/2019.xlsx'
    title_list = ['2000'
        ,'2001XXXXXXXXX'
        ,'2002XXXXXXXXX'
        ,'2003XXXXXXXXX'
        ,'2004XXXXXXXXX'
        ,'2005XXXXXXXXX'
        ,'2006XXXXXXXXX'
        ,'2007XXXXXXXXX'
        ,'2008XXXXXXXXX'
        ,'2009XXXXXXXXX'
        ,'2010XXXXXXXXX'
        ,'2011XXXXXXXXX'
        ,'2012XXXXXXXXX'
        ,'2013XXXXXXXXX'
        ,'2014XXXXXXXXX']
    copyFileto(sourcefile,title_list)













