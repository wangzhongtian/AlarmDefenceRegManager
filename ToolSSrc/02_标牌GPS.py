#-*- coding: UTF-8
#-*- coding: UTF-16
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
		# tplXlsFileName ="基础数据\\reportTpl.xlsx"
		# rootPath = cfgfile.root文件夹
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
		try:
			cls.deltaMax = int( os.environ["deltaMax"].strip() )
		except:
			cls.deltaMax = 0.0
		cls.现场时间调校 =0.0# int( os.environ[ "现场时间调校" ].strip() )

		cls.主机信息文件夹 ="../基础数据";# cls.主机信息文件夹.strip( )

		cls.主机信息文件夹 = cls.主机信息文件夹.strip( )
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

		print( "标牌位置信息文件名")
		TypeIDstr="A001"
		Newestfile,fullpaths,commonPAth= Exdecision.get最新file(TypeIDstr ,cls.主机名号 , cls.root文件夹,exceptFolders)
		cls.标牌坐标文件名 = Newestfile
		cls.SignalDatarootpath = commonPAth
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
def writetuple(Row):
	for k in Row:
		# print( v[-5:-4] )
		# if k in [ "核准经度","归一牌号" ]:
			print(k )

###################################################################
# import CSVProc 
class sd_DictCls (dict):
	def __missing__(self,key) :
		return None

def removeNullKeyinDictArrary( arraryDict1 , keynames = ["通道号","牌号"]):
	cnt = len(arraryDict1 )
	for idx in range(cnt,-1,-1):
		try:
			dict1 = arraryDict1[idx]
		except:
			continue
		isNull =False
		for key in keynames:
			if dict1[ key ] =="" or dict1[ key ] == None:
				isNull =True
		if isNull == True:
			del arraryDict1[idx]
def writeDict(Row):
	# return 
	for k,v in Row.items():
		# print( v[-5:-4] )
		# if k in [ "核准经度","归一牌号" ]:
		# if Row[ "归一牌号" ] [-3:]> "260":
			print(k,v  )
	# print()
	pass

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
# print("Key")
# map(writeDict,KeyGPS_TableDictArrary )	;print();print()
# KeyName: 合并不同Dict时，Keyname相同时，则进行字段的合并
# Arrarys,Fields ：Arrarys dict of dict 的数组，两个数组的个数应该相同。
#fields : for example:[ ("f5") ,("f4","f3") ,("f2" ,"f1")   ]

def CombineDicts_坐标(KeyName, Arrarys,Fields ):
	tgrDict = sd_DictCls()
	datasNum = len(Arrarys )
	for i in range( 0,datasNum ):
		for srcdic in  Arrarys[ i]:
			key =srcdic[ KeyName  ]
			dataDict = tgrDict.setdefault(key, sd_DictCls( {KeyName:key} ) )
			# tgrDict[ key ] = dataDict
			for field in Fields[i]:
				dataDict[ field ] = srcdic.get(field,"")

	return tgrDict

def Combine标牌Dicts(Arrarys):
	合并Key ="归一牌号"
	tgrDict = sd_DictCls()
	for ad in Arrarys:
		for d1 in ad:
			归一牌号= d1[合并Key]
			dd = tgrDict.get( 归一牌号,None )
			if dd ==None:
				tgrDict[ 归一牌号] = d1
			else:
				dd.update( d1 ) 
	return tgrDict

# GenDict = CombineDicts_坐标(KeyName, Arrarys,Fields )
cfgfile.getFiles()

标牌坐标_TableDictArrary = Odbc2ADLib.xlstbl2AD.xls2AD([cfgfile.标牌坐标文件名], ["标牌位置","1通道标牌信息","2通道标牌信息",],SpanLines = 0);print()
KeyGPS_TableDictArrary =  Odbc2ADLib.xlstbl2AD.xls2AD([cfgfile.标牌关键坐标文件名], ["关键点坐标",],SpanLines = 0);print()			
HostInfoDictArrary   = Odbc2ADLib.xlstbl2AD.xls2AD([cfgfile.主机信息文件名], ["5电子围栏信息表（围栏标段填写）"],SpanLines = 2);print() #SpanLines = 2)
HostID = getHostIDFrom主机名号(HostInfoDictArrary,cfgfile.主机名号)

print("-----------1------------")

removeNullKeyinDictArrary( 标牌坐标_TableDictArrary )			
MCls.HostID =  HostID		
map(MCls.归一标牌, 标牌坐标_TableDictArrary  )	
map(MCls.归一标牌, KeyGPS_TableDictArrary )	
KeyName= "归一牌号"
Arrarys=[标牌坐标_TableDictArrary ,KeyGPS_TableDictArrary]
Fields=[("相间距离","岸别") ,( "核准经度","核准纬度","地名")  ]
GenDict =Combine标牌Dicts(Arrarys )

# print("Gendict")
# map(writeDict,GenDict.values() )	;print();print()

# print("标牌坐标_TableDictArrary")
# map(writeDict,标牌坐标_TableDictArrary  );	print()

######################   计算中间点坐标  经度和纬度  ，采用等比例分法确定 ############

# OutAllbiaopaiDictDatas(GenDict ,gfieldNames )
import  界标GPS差分计算   
gpsObj = 界标GPS差分计算.jiebiaoGPSCls(GenDict,HostID) 
print("-----------2------------")
# map( writeDict,GenDict.values() )

tpl = "{:}_{:}_{:0>4}年{:0>2}月{:0>2}日{:0>2}_{:0>2}_{:0>2}.kml"
curDt = datetime.datetime.now()
Prefix="B001-"+cfgfile.主机名号
filenameTpl="_标牌GPS_"
fn1 = tpl.format(Prefix,filenameTpl,curDt.year,curDt.month,curDt.day,curDt.hour,curDt.minute,curDt.second )
fn1=os.path.join( cfgfile.root文件夹 ,fn1)
print( "===",len(GenDict),fn1 )
gpsObj.cal界标坐标(cfgfile.root文件夹 ,fn1 )
# print( len(GenDict) )
Exdecision.is_exed_data(  RIDMax= 100,gongzuoRate =10)
# print( len(GenDict) )
# print("=============")
# map( writeDict,GenDict.values() )
gfieldNames = "归一牌号 相间距离 岸别 核准经度 核准纬度 地名"
xlstgrName = fileProc.tplProc(cfgfile.root文件夹,"G001" ,cfgfile.主机名号,"位置和GPS","../基础数据\\reportTpl.xlsx")
# gfieldNames = "归一牌号 通道号 牌号 岸别 核准经度 核准纬度 地名 测试经度 测试纬度 计算经度 计算纬度"
Odbc2ADLib.xlstblWrite.saveDD2Xls(xlstgrName, "标牌GPS",GenDict ,gfieldNames.split())