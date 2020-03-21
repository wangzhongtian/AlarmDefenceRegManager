#-*- coding: UTF-8
#-*- coding: UTF-16
#####
### 读取X002 防区规划表中的数据，生成标牌地图信息
#################
from __future__ import print_function
import sys
sys.path.append("LibSRC")
print(sys.path)
for i in sys.path:
	print(i)
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

import Exdecision
import sys
import os 
# from datetime import datetime, date, time
import Exdecision
import Odbc2ADLib

import datetime

class fileProc():
	@classmethod
	def getDTFileName(cls,Prefix,filenameTpl):
		tpl = "{:}_{:}_{:0>4}年{:0>2}月{:0>2}日{:0>2}_{:0>2}_{:0>2}.xlsx"
		curDt = datetime.datetime.now()
		filename = tpl.format(Prefix,filenameTpl,curDt.year,curDt.month,curDt.day,curDt.hour,curDt.minute,curDt.second )
		return filename
	@classmethod
	def tplProc(cls,rootPath,PreFix,A2Prefix,A3Prefix,tplXlsFileName="../基础数据\\reportTpl.xlsx"):
		# xlsTplName = os.path.join(rootPath,tplXlsFileName)
		xlsTplName = tplXlsFileName
		tgrFile=cls.getDTFileName(PreFix,A2Prefix+"_"+A3Prefix)
		xlstgrName =os.path.join( rootPath,tgrFile)
		srcfileObj =System.IO.FileInfo(xlsTplName )
		srcfileObj.CopyTo( xlstgrName )
		return xlstgrName
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
		# cls.deltaMax = int( os.environ["deltaMax"].strip() )
		# cls.现场时间调校 = int( os.environ[ "现场时间调校" ].strip() )

		# cls.主机信息文件夹=cls.root文件夹 #r"C:\Cur\定时调测数据"

		# cls.主机信息文件夹 = cls.主机信息文件夹.strip( )
		exceptFolders=["_排除文件夹"]

		print()
		print("防区规划表")
		TypeIDstr="X002"
		Newestfile,fullpaths,commonPAth= Exdecision.get最新file(TypeIDstr ,cls.主机名号 , cls.root文件夹,exceptFolders)
		cls.标牌关键坐标文件名 = Newestfile
		cls.SignalDatarootpath  = commonPAth

import re			


def writeDict(Row):
	for k,v in Row.items():
			print(k,v  )
	pass

cfgfile.getFiles()
class MTransferCls():
	KeyGPS_TableDictArrary = []
	@classmethod
	def Transfer(cls,rowDict):
			d1 = dict()
			d1 [ "归一牌号" ] = rowDict.get("定位界标1" )
			d1 [ "岸别" ] 		= rowDict.get("左右岸" )
			d1 [ "核准经度" ] = rowDict.get("防区起点GPS经度坐标" )
			d1 [ "核准纬度" ] = rowDict.get("防区起点GPS纬度坐标" )
			d1 [ "地名" ] = rowDict.get("地名信息" )
			d1 [ "归一牌号" ] = rowDict.get("定位界标1" )
			cls.KeyGPS_TableDictArrary += [ d1 ]
标牌坐标_TableDictArrary = Odbc2ADLib.xlstbl2AD.xls2AD([cfgfile.标牌关键坐标文件名], ["防区规划","03-防区规划表","1定位型电子围栏防区信息表（围栏标段填写）",],SpanLines = 0);print()

map( MTransferCls.Transfer, 标牌坐标_TableDictArrary )
KeyGPS_TableDictArrary = MTransferCls.KeyGPS_TableDictArrary
print("-----------1------------")


tpl = "{:}_{:}_{:0>4}年{:0>2}月{:0>2}日{:0>2}_{:0>2}_{:0>2}.kml"
curDt = datetime.datetime.now()
Prefix="B001-"+cfgfile.主机名号
filenameTpl="_标牌GPS_"
fn1 = tpl.format(Prefix,filenameTpl,curDt.year,curDt.month,curDt.day,curDt.hour,curDt.minute,curDt.second )
fn1=os.path.join( cfgfile.root文件夹 ,fn1)


Exdecision.is_exed_data(  RIDMax= 100,gongzuoRate =10)
import datetime 
import  KeyGPS处理Cls   
gpsObj = KeyGPS处理Cls.keyGPSroccls(KeyGPS_TableDictArrary) 

tpl = "{:}_{:}_{:0>4}年{:0>2}月{:0>2}日{:0>2}_{:0>2}_{:0>2}"
curDt = datetime.datetime.now()
Prefix="B002-"+cfgfile.主机名号
filenameTpl="_KeyGPS_"
fn1 = tpl.format(Prefix,filenameTpl,curDt.year,curDt.month,curDt.day,curDt.hour,curDt.minute,curDt.second )
gpsObj.show界标坐标(cfgfile.SignalDatarootpath ,fn1+".kml")
Exdecision.is_exed_data(  RIDMax= 100,gongzuoRate =10)