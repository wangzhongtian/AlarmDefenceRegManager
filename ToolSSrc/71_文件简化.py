# -*- coding: UTF-8
from __future__ import print_function
import clr
import System
import System.IO

def getFullNames(dllname):
    def getEnvs(envnmame):
        tgt1=System.EnvironmentVariableTarget.Machine
        tgt2=System.EnvironmentVariableTarget.User
        tgt3=System.EnvironmentVariableTarget.Process
        Libpa =[]
        for tgt in (tgt3,tgt1,tgt2):
            Libpaths = System.Environment.GetEnvironmentVariable(envnmame,tgt)
            #print(Libpaths,tgt );print()
            if Libpaths  == None:
                pass
            else:
                Libpa += Libpaths.split(";")
        print()
        return Libpa
    libpaths = [".\\"]+getEnvs("libpath") + getEnvs("path")
    #print( "---------------",libpaths);print()
    for p in libpaths:#.split(";"):
        filename = p+"\\"+dllname
        if ( System.IO.File.Exists(filename) ):
            print("---Load Dlls:" ,filename)
            return filename
    print( "Can not find Lib:" ,dllname)
    return None

dllname=r"""STDLIBalls.DLL""" ;clr.AddReferenceToFileAndPath( getFullNames(dllname) )
dllname=r"""AppAssembly.dll""" ;clr.AddReferenceToFileAndPath( getFullNames(dllname) )

import os
# import datetime
#from datetime import datetime, date, time
import Exdecision
import XLSCombine
import Exdecision
import Odbc2ADLib
import glob
class cfgfile():
	@classmethod
	def getFiles(cls):

		cls.rootpath = r""" E:\ipy\电子围栏Tools\tools\发布防区表\xxxx """.strip() 
		cls.xlsFiles +=glob.glob( os.path.join(cls.rootpath ,"*.xls?") )
##################################################################
def writeDict(Row):
	# return 
	for k,v in Row.items():
		# print( v[-5:-4] )
		# if k in [ "核准经度","归一牌号" ]:
			print(k,v  )
	print()
	pass

def 精简处理(row):
# 左右岸 界标之间围栏长度 
	IDstr= row.get("定位界标1","xxxxxxxxxx")
	row["主机编号"] = "a"+IDstr[1:10]
	row["通道号"] = "a"+IDstr[10:11]
	row["标牌号"] = "a"+IDstr[11:14]
	row["标牌光程"] = row.get("定位界标1光程","无") 
	row["经度坐标"] = row.get("防区起点GPS经度坐标","无") 
	row["纬度坐标"] = row.get("防区起点GPS纬度坐标","无") 
	row["标牌位置信息"] = row.get("地名信息","无") 
	#row["标牌位置信息"] = row.get("界标之间围栏长度","xx") 
	pass

class fileProc():
	@classmethod
	def getDTFileName(cls,Prefix,filenameTpl):
		tpl = "{:}_{:}_{:0>4}年{:0>2}月{:0>2}日{:0>2}_{:0>2}_{:0>2}.xlsx"
		curDt = datetime.now()
		filename = tpl.format(Prefix,filenameTpl,curDt.year,curDt.month,curDt.day,curDt.hour,curDt.minute,curDt.second )
		return filename
	@classmethod
	def tplProc(cls,rootPath,PreFix,A2Prefix,A3Prefix,tplXlsFileName="基础数据\\reportTpl.xlsx"):
		# tplXlsFileName ="基础数据\\reportTpl.xlsx"
		# rootPath = cfgfile.root文件夹
		xlsTplName = os.path.join(rootPath,tplXlsFileName)
		tgrFile=cls.getDTFileName(PreFix,A2Prefix+"_"+A3Prefix)
		xlstgrName =os.path.join( rootPath,tgrFile)
		srcfileObj =System.IO.FileInfo(xlsTplName )
		srcfileObj.CopyTo( xlstgrName )
		return xlstgrName
import re 
def mainentry( filename ):
	print("Process the  files:" ,filename)
	SignalRowNAme ="定位界标1"
	SpanLineNum =0
	while True:
		#print( SpanLineNum)
		if SpanLineNum >5 :
			print("Excel file Format Error ,Spanlines exceed 5 row.")
			return 
		try:
			ad防区数据= Odbc2ADLib.xlstbl2AD.xls2AD( [ filename ], ["Sheet2","Sheet1","03-防区规划表","1定位型电子围栏防区信息表（围栏标段填写）"],SpanLines =SpanLineNum,MaxRecords=3000);print()
		except:
			print("Excel file Format Error ,too few rows.")
			return
		if len(ad防区数据 ) <= 1:
			SpanLineNum +=1
			continue
		if  SignalRowNAme in ad防区数据[0].keys():
			#map( writeDict ,ad防区数据 )
			break
		else:
			SpanLineNum +=1
	map( 精简处理 ,ad防区数据 )
	#print(""----------------"");print()
	#map( writeDict ,ad防区数据 )
	#return
	gFieldName="主机编号 通道号	标牌号	标牌光程 经度坐标 纬度坐标 标牌位置信息 左右岸 界标之间围栏长度 "

	#xlstgrName
	dirname,fn1 =os.path.split( filename)
	if re.match("^A300_标牌信息记录_",fn1) != None:
		return 
	#fn1="1定位型电子围栏防区信息表-郑州002号主机-精调.xlsx"
	m1 = re.match("[_-]*[^-_]+[-_]+([^0-9]+[0-9]{3})号主机*",fn1)
	主机名号= ""
	if  m1 != None :
		主机名号=m1.group(1)
		print( 主机名号);print()
	else:
		print( "文件名称规则不符，未处理",fn1 )
		return
	#return 
	tgrname="A300_标牌信息记录_"+ 主机名号+"号主机.xlsx"
	xlstgrName = os.path.join(dirname ,tgrname )

	xlsTplName=  os.path.join(dirname ,"reportTpl.xlsx" )
	srcfileObj =System.IO.FileInfo(xlsTplName )
	srcfileObj.CopyTo( xlstgrName )
	#ad_data = 现场调测整理(dd_现场数据 )
	Odbc2ADLib.xlstblWrite.AD2Xls(xlstgrName, "标牌数据记录和处理结果",ad防区数据,gFieldName.split())

Exdecision.is_exed_data(  RIDMax= 100,gongzuoRate =10)
cfgfile.getFiles()
for xlc in cfgfile.xlsFiles:
	print(xlc)
	mainentry(xlc)
	#break