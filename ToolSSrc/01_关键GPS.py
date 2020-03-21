#-*- coding: UTF-8
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
    libpaths = ["./"]+getEnvs("libpath") + getEnvs("path")
    print( "---------------",libpaths);print()
    for p in libpaths:#.split(";"):
        filename = p+"/"+dllname
        if ( System.IO.File.Exists(filename) ):
            print("---Load Dlls:" ,filename)
            return filename
    print( "Can not find Lib:" ,dllname)
    return None

dllname=r"""StdLibALLs.dll""" ;clr.AddReferenceToFileAndPath( getFullNames(dllname) )
dllname=r"""AppAssembly.dll""" ;clr.AddReferenceToFileAndPath( getFullNames(dllname) )

import sys
# sys.path=[]
# sys.path.append("LibSRC")
# print(sys.path)
# for i in sys.path:
# 	print(i,"  ---")
# print("----------")

import os 
from datetime import datetime, date, time
# import Exdecision

import Odbc2ADLib
import Exdecision
class cfgfile():
	@classmethod
	def getFiles(cls):
		try:
			if 主机名号=="":
				cls.主机名号   = os.environ["主机名号"].strip() #"磁县002号"
			else:
				cls.主机名号   = 主机名号.strip() #"磁县002号"
		except:
			cls.主机名号   = os.environ["主机名号"].strip() #"磁县002号"
		# cls.root文件夹 = os.environ["root文件夹"].strip() 
		cls.root文件夹 ="../定时调测数据"
		try:
			cls.deltaMax = int( os.environ["deltaMax"].strip() )
		except:
			cls.deltaMax = 0.0
		cls.现场时间调校 =0.0# int( os.environ[ "现场时间调校" ].strip() )

		cls.主机信息文件夹=cls.root文件夹 #r"C:\Cur\定时调测数据"

		cls.主机信息文件夹 ="../基础数据";# cls.主机信息文件夹.strip( )
		exceptFolders=["_排除文件夹"]

		# for idx in range(0,len(exceptFolders) ):
		# 	exceptFolders[idx] = exceptFolders[idx].lower()

		for folder in exceptFolders:
			print( folder)
		print();print( "主机信息文件名")
		TypeIDstr="A100"
		Newestfile,fullpaths,commonPAth= Exdecision.get最新file(TypeIDstr ,None ,  cls.主机信息文件夹,exceptFolders)
		cls.主机信息文件名 = Newestfile

		print()
		print("标牌关键坐标文件名")
		TypeIDstr="A002"
		Newestfile,fullpaths,commonPAth= Exdecision.get最新file(TypeIDstr ,cls.主机名号 , cls.root文件夹,exceptFolders)
		cls.标牌关键坐标文件名 = Newestfile
		cls.SignalDatarootpath  = commonPAth
import re			
def getIPFrom主机名号( _HostInfoDictArrary,主机名号 ):
	for row in _HostInfoDictArrary:
		host = row["管理处名称"] 
		if re.match( "^"+host,主机名号):
			IPStr = row["IP"]
			return IPStr	
def getHostIDFrom主机名号(_HostInfoDictArrary,主机名号):
	for row in _HostInfoDictArrary:
		host = row["管理处名称"] 
		if re.match( "^"+host,主机名号):
			IPStr = row["电子围栏主机编号"]
			return IPStr[1:]
###################################################################
# import CSVProc 
class sd_DictCls (dict):
	def __missing__(self,key) :
		return None

def 归一化标牌计算(dictArrary):
	for rowDict in dictArrary:
		# newValue =""
		# if "归一牌号" in rowDict.keys():
		# 	continue
		value1 = rowDict[ "通道号" ]
		value2 = rowDict[ "牌号" ]
		value = "_".join( (value1[-2:],value2[-3:]) )
		# print(value,value1,value2 )
		rowDict [ "归一牌号"] = value

cfgfile.getFiles()
HostInfoDictArrary   = Odbc2ADLib.xlstbl2AD.xls2AD([cfgfile.主机信息文件名], ["5电子围栏信息表（围栏标段填写）"],SpanLines = 2);print() #SpanLines = 2)
HostID = getHostIDFrom主机名号(HostInfoDictArrary,cfgfile.主机名号)
KeyGPS_TableDictArrary   = Odbc2ADLib.xlstbl2AD.xls2AD([cfgfile.标牌关键坐标文件名], ["关键点坐标"],SpanLines = 0);print() #SpanLines = 2)

class MCls():
	HostID = ""
	@classmethod
	def 归一标牌(cls,rowDict):
			value1 =  rowDict.get("通道号" )
			value2 = rowDict.get( "牌号")
			if value1 == None or value2 == None:
				return 
			value = "_".join( (value1[-1:],value2[-3:]) )
			# print(value )
			rowDict [ "归一牌号"] = cls.HostID + value
MCls.HostID =  HostID	
map(MCls.归一标牌, KeyGPS_TableDictArrary )	
######################   计算中间点坐标  经度和纬度  ，采用等比例分法确定 ############
import  KeyGPS处理Cls   
gpsObj = KeyGPS处理Cls.keyGPSroccls(KeyGPS_TableDictArrary) 

tpl = "{:}_{:}_{:0>4}年{:0>2}月{:0>2}日{:0>2}_{:0>2}_{:0>2}"
curDt = datetime.now()
Prefix="B002-"+cfgfile.主机名号
filenameTpl="_KeyGPS_"
fn1 = tpl.format(Prefix,filenameTpl,curDt.year,curDt.month,curDt.day,curDt.hour,curDt.minute,curDt.second )
gpsObj.show界标坐标(cfgfile.SignalDatarootpath ,fn1+".kml")
Exdecision.is_exed_data(  RIDMax= 100,gongzuoRate =10)
# OutAllbiaopaiKeyGPSDict(KeyGPS_TableDictArrary ,gfieldNames )
import datetime
class fileProc():
	@classmethod
	def getDTFileName(cls,Prefix,filenameTpl):
		tpl = "{:}_{:}_{:0>4}年{:0>2}月{:0>2}日{:0>2}_{:0>2}_{:0>2}.xlsx"
		curDt = datetime.datetime.now()
		filename = tpl.format(Prefix,filenameTpl,curDt.year,curDt.month,curDt.day,curDt.hour,curDt.minute,curDt.second )
		return filename
	@classmethod
	def tplProc(cls,rootPath,PreFix,A2Prefix,A3Prefix,tplXlsFileName="基础数据\\reportTpl.xlsx"):
		# tplXlsFileName ="基础数据\\reportTpl.xlsx"
		# rootPath = cfgfile.root文件夹
		# xlsTplName = os.path.join(rootPath,tplXlsFileName)
		xlsTplName = tplXlsFileName
		tgrFile=cls.getDTFileName(PreFix,A2Prefix+"_"+A3Prefix)
		xlstgrName =os.path.join( rootPath,tgrFile)
		srcfileObj =System.IO.FileInfo(xlsTplName )
		srcfileObj.CopyTo( xlstgrName )
		return xlstgrName

xlstgrName = fileProc.tplProc(cfgfile.root文件夹,"A002" ,cfgfile.主机名号,"KeyGPS","../基础数据\\reportTpl.xlsx")

# gfieldNames = "归一牌号 通道号 牌号 岸别 核准经度 核准纬度 地名 测试经度 测试纬度 计算经度 计算纬度"
Odbc2ADLib.xlstblWrite.AD2Xls(xlstgrName, "关键点坐标",KeyGPS_TableDictArrary)#,gfieldNames.split())

