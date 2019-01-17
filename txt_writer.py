#!/usr/bin/env python
# _*_ UTF-8 _*_
#双层列表写入文件
#第一种方法，每一项用空格隔开，一个列表是一行写入文件
data =[ ['a','b','c'],['a','b','c'],['a','b','c']]
with open("test.txt","w") as f: #设置文件对象
    for i in data:  #对于双层列表中的数据
        i = str(i).strip('[').strip(']').replace(',','').replace('\'','')+'\n'
        #将其中每一个列表规范化成字符串
        f.write(i)
        #写入文件
#第二种方法，直接将每一项都写入文件
data =[ ['a','b','c'],['a','b','c'],['a','b','c']]
with open("test.txt","w") as f:
    #设置文件对象
    for i in data:
        #对于双层列表中的数据
        f.writelines(i)
        #写入文件
#第三种方法 将数组写入文件
import numpy as np
#第一种方法
np.savetxt("data.txt",data)     #将数组中数据写入到data.txt文件
#第二种方法
np.save("data.txt",data)        #将数组中数据写入到data.txt文件












