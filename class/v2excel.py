import os
from posixpath import split    #导入os库，这个库可以实现我们要的功能，这个库是官方的，不用另外下载
import re    #导入os库，这个库可以实现我们要的功能，这个库是官方的，不用另外下载
import xlwt
import openpyxl
import codecs
from openpyxl.utils import get_column_letter
import sys
import shutil

shutil.rmtree('EXCE/')
os.mkdir('EXCE/')

shutil.rmtree('EXU2_POS/')
os.mkdir('EXU2_POS/')

def verilogpos(verilogfile,posverilogfile):

    f = open(verilogfile,'r')
    p = open(posverilogfile,'w')
     
    for line in f.readlines():
        line0 = line.replace(' ','')
        vld_pos = line0.find("//");
        stp_pos = line0.find(",")
        if vld_pos>=0:
            line1 = line0[0:vld_pos]
        else:
            if (stp_pos < 0):
                line1 = line0[0:vld_pos]
            else:
                line1 = line0[0:stp_pos+1]
        splits0 = line1.find("put")
        splits1 = line1.find("]")
        splits2 = line1.find(",")
        if splits0>0 and splits1 > 0 and splits2 >0:
            line3 = line1[splits1+1:splits2+1] + line1[splits0+3:splits1+1]+ ',' +',' + line1[0:splits0+3] + ',' + '\n'
        elif splits0>0 and splits2 >0:
            line3 = line1[splits0+3:splits2+1] + '[0:0]' + ',' + ',' + line1[0:splits0+3] + ',' + '\n'
        elif splits0>0:
            line3 = line1[splits0+3:] + ',' + '[0:0]' + ',' + ',' + line1[0:splits0+3] + ',' + '\n'
        else:
            line3 = '' 
        p.writelines(line3)
    f.close()
    p.close()


def verilog2xls(filename,xlsname):
#文本转换成xls的函数
#param filename txt文本文件名称、
#param xlsname 表示转换后的excel文件名
    try:
        f = open(filename)
        xls=xlwt.Workbook()
        #生成excel的方法，声明excel
        sheet = xls.add_sheet('sheet1',cell_overwrite_ok=True)
        x = 0
        while True:
            #按行循环，读取文本文件
            line = f.readline()
            if not line:
                break  #如果没有内容，则退出循环
            for i in range(len(line.split(','))):
                item=line.split(',')[i]
                sheet.write(x,i,item) #x单元格经度，i 单元格纬度
            x += 1 #excel另起一行
        f.close()
        xls.save(xlsname) #保存xls文件
    except:
        raise
if __name__ == "__main__" :
    filename = os.listdir('EXU2')
    xlsname  = os.listdir('EXU2')
    posverilogfile  = os.listdir('EXU2')
    for i in range(len(filename)):    #构造文件路径
        filename[i] = 'EXU2' + '/' + filename[i]
        posverilogfile[i] = 'EXU2_POS' + '/' + posverilogfile[i]
        xlsname[i]  = 'EXCE' + '/' + xlsname[i] + ".xlsx"
        verilogpos(filename[i],posverilogfile[i])
        verilog2xls(posverilogfile[i],xlsname[i])