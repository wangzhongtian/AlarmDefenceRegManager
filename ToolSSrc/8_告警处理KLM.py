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
    libpaths = [".\\"]+getEnvs("libpath") + getEnvs("path")
    #print( "---------------",libpaths);print()
    for p in libpaths:#.split(";"):
        filename = p+"\\"+dllname
        if ( System.IO.File.Exists(filename) ):
            print("---Load Dlls:" ,filename)
            return filename
    print( "Can not find Lib:" ,dllname)
    return None


dllname=r"""STDLIBAlls.DLL""" ;clr.AddReferenceToFileAndPath( getFullNames(dllname) )
dllname=r"""AppAssembly.DLL""" ;clr.AddReferenceToFileAndPath( getFullNames(dllname) )
import collections
import Exdecision
import os
from datetime import datetime, date, time
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
		try:
			cls.deltaMax = int( os.environ["deltaMax"].strip() )
		except:
			cls.deltaMax = 0.0
		cls.现场时间调校 =0.0# int( os.environ[ "现场时间调校" ].strip() )

		return
		cls.主机信息文件夹 ="../基础数据";# cls.主机信息文件夹.strip( )
		exceptFolders=["_排除文件夹"]

		# for idx in range(0,len(exceptFolders) ):
		# 	exceptFolders[idx] = exceptFolders[idx].lower()

		for folder in exceptFolders:
			print( folder)
		cls.主机信息文件夹 = cls.主机信息文件夹.strip( )
		print();print( "标牌坐标文件名")
		TypeIDstr="G001"
		Newestfile,fullpaths,commonPAth= Exdecision.get最新file(TypeIDstr ,cls.主机名号 ,cls.root文件夹,exceptFolders)
		cls.标牌坐标文件名 = Newestfile
		cls.SignalDatarootpath = commonPAth

		print();print ("标牌光程文件名")
		TypeIDstr="M003"
		Newestfile,fullpaths,commonPAth= Exdecision.get最新file(TypeIDstr ,cls.主机名号 ,  cls.root文件夹,exceptFolders)
		cls.标牌光程文件名 = Newestfile

		print();print( "主机信息文件名")
		TypeIDstr="A100"
		Newestfile,fullpaths,commonPAth= Exdecision.get最新file(TypeIDstr ,None ,  cls.主机信息文件夹,exceptFolders)
		cls.主机信息文件名 = Newestfile

		print();print( "标牌归属文件名")
		TypeIDstr="A003"
		Newestfile,fullpaths,commonPAth= Exdecision.get最新file(TypeIDstr ,cls.主机名号 , cls.root文件夹,exceptFolders)
		cls.标牌归属文件名 = Newestfile

###################################################################

class 光程DictCls (dict):
	def __missing__(self,key) :
		return None
class cntCls():
	countCls = collections.Counter()
	@classmethod
	def updateCnt(cls, GlFangquID, Alarm_dt):
		cls.countCls[ (GlFangquID,Alarm_dt.hour )] +=1

		
import datetime
#告警号（类型）	告警时间	最后的时间	主机号	防区ID
def 振动入侵告警筛选(row):
	d={}
	aID = row.get("告警号（类型）",None)
	aID=int(float(aID))
	# print("---", aID )
	if aID in [101,301,401,"101","301","401"]:
		dt = row.get("告警时间",None)
		format1="yyyy年M月d日  H:m:s"
		format1="%Y年%m月%d日 %H:%M:%S" 
		d1 = datetime.datetime.strptime( dt,format1)#"%Y年%m月%d日  %H:%M:%s")
		# yyyy'/'M'/'d' 'H':'m':''

		# "%d/%m/%y %H:%M
		d[ "告警时间"] = d1
		# machineID = int(float(row.get("主机号",None) ))
		# d[ "主机号"] =machineID 
		fangqu = row.get("防区名称",None)
		try:
			regionID = int(float(fangqu ))
		except:
			regionID=fangqu
		# d[ "防区名称"] = regionID
		# d[ "告警号（类型）"] = aID
		cntCls.updateCnt(regionID,d1)
		# return d
	return None


cfgfile.getFiles()
print()
def writeDict(row):
	for key,val in row.items():
		print( key,"--",val)
	print()
# print(cfgfile.标牌光程文件名)


# HostInfoDictArrary   = Odbc2ADLib.xlstbl2AD.xls2AD([cfgfile.主机信息文件名], ["5电子围栏信息表（围栏标段填写）"],SpanLines = 2);print() #SpanLines = 2)



	
def getHostIDFromSigID(_GuiyiPaihao):
	if len( _GuiyiPaihao ) <9:
	   return None
	HostIDstr = None
	HostIDstr = _GuiyiPaihao[:5]+"0"+_GuiyiPaihao[6:9]
	# print( HostIDstr )
	return HostIDstr

def getIPFromHostInfoTbl( _HostInfoDictArrary,HostIDSt ):
	guiyiHostID = getHostIDFromSigID(HostIDSt )
	for row in _HostInfoDictArrary:
		host = row["电子围栏主机编号"]
		# IP = row["IP地址"]
		# 主机友好名 = row["主机友好名"] 
		if guiyiHostID == getHostIDFromSigID(host ):
			IPStr = row["IP"]
			return IPStr

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
			return IPStr

OrgfieldNames="归一牌号	归一牌号	IP	通道号(归一牌号[:10])	子防区号(顺序编号)	拟合光程	拟合光程	经度	纬度	经度	纬度	岸别	地名	相间距离 归属部门"
RegionFieldNames="定位界标1	定位界标2	电子围栏主机IP	通道号	子防区号	定位界标1光程	定位界标2光程	防区起点GPS经度坐标	防区起点GPS纬度坐标	防区终点GPS经度坐标	防区终点GPS纬度坐标	左右岸	地名信息	界标之间围栏长度 备注"

def getHostID(_GenDict):
	for key in _GenDict.keys():
		return key[:9]
	return None
import collections

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
		xlsTplName =tplXlsFileName;# os.path.join(rootPath,tplXlsFileName)
		tgrFile=cls.getDTFileName(PreFix,A2Prefix+"_"+A3Prefix)
		xlstgrName =os.path.join( rootPath,tgrFile)
		srcfileObj =System.IO.FileInfo(xlsTplName )
		srcfileObj.CopyTo( xlstgrName )
		return xlstgrName

filename= r"""F:\\ipy\\电子围栏Tools\\ToolSSrc\\防区告警信息表20170103_20170118.xls"""
AlarmLog_ad  = Odbc2ADLib.xlstbl2AD.xls2AD( [ filename ], ["Sheet"]);print()

AlarmsAD = filter(None,map( 振动入侵告警筛选,AlarmLog_ad ) )

#101 301 401
# filter( writeDict,AlarmsAD )
print( "振动光缆入侵{},告警总数{}".format( len(AlarmsAD ) , len(AlarmLog_ad )) )
统计表=[]
for i in cntCls.countCls.most_common():
	d1 ={ "RegionID" :i[0][0],"TimeArea":i[0][1],"告警数量": i[1]  }
 # RegionID : 
	# print(,i[0][1],AlarMCnt:,i[1])
	统计表.append(d1 )
# cfgfile.root文件夹=""
# cfgfile.主机名号 = "---"

xlstgrName = fileProc.tplProc(cfgfile.root文件夹,"G002" ,cfgfile.主机名号,"告警统计","../基础数据\\reportTpl.xlsx")
Odbc2ADLib.xlstblWrite.AD2Xls(  xlstgrName,"统计报表",统计表	) #,RegionFieldNames.split()
print( xlstgrName)
