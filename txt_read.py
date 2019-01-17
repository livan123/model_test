#!/usr/bin/env python
# _*_ UTF-8 _*_
#第一种方法
f = open("test.txt","r")   #设置文件对象
line = f.readline()
line = line[:-1]
while line:             #直到读取完文件
    line = f.readline()  #读取一行文件，包括换行符
    line = line[:-1]     #去掉换行符，也可以不去
f.close() #关闭文件
#第二种方法
data = []
for line in open("test.txt","r"): #设置文件对象并读取每一行文件
    data.append(line)               #将每一行文件加入到list中
#第三种方法
f = open("test.txt","r")   #设置文件对象
data = f.readlines()  #直接将文件中按行读到list里，效果与方法2一样
f.close()             #关闭文件
# 第四种方法：
import numpy as np
data = np.loadtxt("test.txt")   #将文件中数据加载到data数组里












