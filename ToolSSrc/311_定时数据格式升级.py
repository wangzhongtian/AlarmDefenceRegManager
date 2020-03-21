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
from datetime import datetime, date, time
# from cfgfile import *
import Exdecision
import XLSCombine
import Exdecision
import Odbc2ADLib
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
		cls.root文件夹 = os.environ["root文件夹"].strip() 
		cls.deltaMax = int( os.environ["deltaMax"].strip() )
		cls.现场时间调校 =0.0# int( os.environ[ "现场时间调校" ].strip() )

		cls.主机信息文件夹=cls.root文件夹 #r"C:\Cur\定时调测数据"

		cls.主机信息文件夹 ="基础数据";# cls.主机信息文件夹.strip( )
		cls.srcAlmfilenames=[]
		cls.SignalDatarootpath = cls.root文件夹
		cls.srcTriggerfilenames =[]
		cls.rootpath = cls.root文件夹
		exceptFolders=["_排除文件夹"]

		for folder in exceptFolders:
			print( folder)

		print ("告警日志文件")
		try:
			TypeIDstr="A007"#"A005"
			Newestfile,fullpaths,commonPAth= Exdecision.get最新file(TypeIDstr ,cls.主机名号, cls.root文件夹,exceptFolders)
			# print( "============",fullpaths)
			if fullpaths != None:
				cls.srcAlmfilenames = cls.srcAlmfilenames +fullpaths
				cls.SignalDatarootpath = commonPAth
		except:
			pass
		# try:
		# 	print()
		# 	print ("触发记录文件")
		# 	TypeIDstr="A004"#"A004"
		# 	Newestfile,fullpaths,commonPAth= Exdecision.get最新file(TypeIDstr ,cls.主机名号 , cls.root文件夹,exceptFolders)
		# 	if fullpaths != None:
		# 		cls.srcTriggerfilenames = cls.srcTriggerfilenames + fullpaths
		# 	cls.rootpath = commonPAth	
		# except:
		# 	pass
	
###################################################################

def getsecondslen( curdateStr,curTimeStr ):
	relDt =datetime(2016, 7, 28)
	strtime = curdateStr  + " " + curTimeStr
	l = 0.0
	try:
		strtime = strtime.replace(";",":").replace("；",":")
		t=datetime.strptime(strtime,"%Y/%m/%d %H:%M:%S")
		l= (t-relDt).total_seconds()
	except:
		print("Error:strtime is :",strtime)
		return -1000
	return l

def getTime(日期Str,时间1Str):
		datestr= 日期Str
		timestr = 时间1Str	
		timestr=时间1Str.replace("：",":").replace("；",":").replace(" ","")
		try:
			l= getsecondslen( datestr,timestr)+cfgfile.现场时间调校
			return l
		except:
			return -10000

def writeDict(Row):
	# return 
	for k,v in Row.items():
		# print( v[-5:-4] )
		# if k in [ "核准经度","归一牌号" ]:
			print(k,v  )
	print()
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


def 现场调测整理(dd_现场数据):
	import collections
	调测结果dict=dict( ) #dict of tiaoce shuju ,key is 归一牌号
	paihao =""
	for begTime  in sorted( dd_现场数据.keys() ,reverse = True):
		row = dd_现场数据[begTime ]
		paihao= row["归一牌号" ]
		if paihao not in 调测结果dict:
			调测结果dict[ paihao ] = row


	resAD = [] 
	for paihao in sorted( 调测结果dict.keys() ):
		row = 调测结果dict[ paihao ]
		resAD += [row]

	return resAD


def 时间区间分析法():
	print("Process the  files:")
	# cfgfile.getFiles() 
	ad现场数据 = Odbc2ADLib.xlstbl2AD.xls2AD( cfgfile.srcTriggerfilenames, ["2通道","1通道","标牌敲击测试记录表"],SpanLines = 1,MaxRecords=3000);print()
	# print( len(ad现场数据 ))
	# map( writeDict,ad现场数据)
	# print( len(ad告警数据 ))
	# map( writeDict,ad告警数据)
	ad_data = 现场调测整理(dd_现场数据 )
	map( 显示处理,dd_现场数据.values())
	# print("=====================")
	# map( writeDict,dd_alarmData.values() )

	xlstgrName = fileProc.tplProc(cfgfile.root文件夹,"G007" ,cfgfile.主机名号,"测试数据","基础数据\\reportTpl.xlsx")
	gFieldName="归一牌号	通道号	标牌号	起始敲击时间	最后一次敲击结束时间	触发结束时间	触发开始时间	光程均值	首选均值	次选均值	三选均值	首选邻值	次选邻值	三选邻值	敲击次数	日期	主机时间减去工人手表时间差"


	# map( writeDict,ad_data )
	Odbc2ADLib.xlstblWrite.AD2Xls(xlstgrName, "光程调试结果",ad_data ,gFieldName.split())
def entry定时同步数据处理():
	cfgfile.getFiles()
	Exdecision.is_exed_data(  RIDMax= 100,gongzuoRate =10)
	时间区间分析法()

entry定时同步数据处理()