#-*- coding: UTF-8
from __future__ import print_function

import clr
import System
import System.IO
import sys
sys.path.append("LibSRC")
print(sys.path)
for i in sys.path:
	print(i)
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
		# cls.root文件夹 = os.environ["root文件夹"].strip() 
		cls.root文件夹 ="../定时调测数据"
		try:
			cls.deltaMax = int( os.environ["deltaMax"].strip() )
		except:
			cls.deltaMax = 0.0
		cls.现场时间调校 =0.0# int( os.environ[ "现场时间调校" ].strip() )

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



cfgfile.getFiles()
print()

def writeDict(row):
	for key,val in row.items():
		print( key,"--",val)
	print()
# print(cfgfile.标牌光程文件名)
Fl_TableDictArrary   = Odbc2ADLib.xlstbl2AD.xls2AD( [ cfgfile.标牌光程文件名 ], ["02-界标光程表","界标光程表"]);print()
Gps_TableDictArrary  = Odbc2ADLib.xlstbl2AD.xls2AD([cfgfile.标牌坐标文件名], ["01-界标坐标表","标牌GPS"]);print()
Host_TableDictArrary = Odbc2ADLib.xlstbl2AD.xls2AD([cfgfile.标牌归属文件名], ["界标归属","A1通道标牌信息","A2通道标牌信息","03-界标归属","1通道标牌信息","2通道标牌信息"]);print()

# map(writeDict,Host_TableDictArrary )
HostInfoDictArrary   = Odbc2ADLib.xlstbl2AD.xls2AD([cfgfile.主机信息文件名], ["5电子围栏信息表（围栏标段填写）"],SpanLines = 2);print() #SpanLines = 2)
# map(writeDict,HostInfoDictArrary )

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

# KeyName: 合并不同Dict时，Keyname相同时，则进行字段的合并
# Arrarys,Fields ：Arrarys dict of dict 的数组，两个数组的个数应该相同。
#fields : for example:[ ("f5") ,("f4","f3") ,("f2" ,"f1")   ]
def CombineDicts(KeyName, Arrarys,Fields ):
	tgrDict = 光程DictCls()
	datasNum = len(Arrarys )
	for i in range( 0,datasNum ):
		fields = Fields[i]
# 		print( fields )
		arraryDict = Arrarys[ i]
		for srcdic in arraryDict:
			# print( i )
			key =srcdic.get(KeyName ,None)
			if key == None:
				continue
			else:
				dataDict = tgrDict.get(key,None)
			if dataDict ==None:
				dataDict= 光程DictCls()
			# print(key,i)
			dataDict[ KeyName ] = key
			for field in fields:
				# print( field )
				dataDict[ field ] = srcdic.get( field ,"")
				dataDict[ "IP" ] =  getIPFromHostInfoTbl( HostInfoDictArrary,key )
			tgrDict[ key ] = dataDict
	return tgrDict
def outarraryofDict( arraryDict):
	for row in arraryDict:
		for k in row.keys():
			print( "{:}={:}".format(k, row[k]) )

# outarraryofDict(Gps_TableDictArrary)


def gen归一牌号( _TableDictArrary):
	for row in _TableDictArrary:
		# print( len(row))
		keys = row.keys()
		for k in keys:
			print( k,end="," )
		print()
		if "归一牌号" in keys:
			# print( "0000")
			continue
		elif "通道号" in keys and "牌号" in keys  :
			通道号 = row["通道号"]
			牌号 = row["牌号"]
			归一牌号 = 通道号[ -10:]+"_"+ 牌号[-3:]
			# print(归一牌号,通道号 ,牌号 )
			row[ "归一牌号" ] = 归一牌号
		else:
			print( "Nothing in dict keys ")

def 归一牌号(row):
	keys = row.keys()
	if "归一牌号" in keys:
		# print( "0000")
		return
	elif "通道号" in keys and "牌号" in keys  :
		通道号 = row["通道号"]
		牌号 = row["牌号"]
		归一牌号 = ghostID+通道号[ -1:]+"_"+ 牌号[-3:]
		# print(归一牌号,通道号 ,牌号 )
		row[ "归一牌号" ] = 归一牌号
	else:
		pass
ghostID = getHostIDFrom主机名号(HostInfoDictArrary ,cfgfile.主机名号)
ghostID =ghostID[1:]
		#print( "Nothing in dict keys ")
print(" ---------------------------")
map(归一牌号, Fl_TableDictArrary)
map(归一牌号, Gps_TableDictArrary)
map(归一牌号, Host_TableDictArrary)

print(" ---------------------------")
KeyName= "归一牌号"
Arrarys=[Fl_TableDictArrary ,Gps_TableDictArrary, Host_TableDictArrary]
Fields=[( "拟合光程",),("核准经度","核准纬度","相间距离","岸别","地名",),("归属部门",)   ] #"GPS测试",
GenDict = CombineDicts(KeyName, Arrarys,Fields )



def 显示归一( row):
	fval= row["拟合光程"]
	# print( fval)
	if fval != None:
		val =float(fval)
	else:
		val = 0
	row["拟合光程"] = int(val)

	# val = row["地名"] 
	# print( "-"+val+"+" )
	# if 	val == "":
	# 	row["地名"] = None

map( 显示归一, GenDict.values())
# map(writeDict,GenDict.values() )
# for row in GenDict.values():
# 	显示归一( row)
def OutAllbiaopaiDatas(_GenDict ,_fieldNames  ):	
	CsvObj = CSVProc.CSVProcCls()
	Ids= sorted(_GenDict.keys() )
 	for sID in Ids:
		RowDict = _GenDict[ sID ]
# 		RowDict[ "归一牌号" ] = g主机编号 + sID
	CsvObj.outDict2Xls(_GenDict,_fieldNames,"防区规划_",cfgfile.SignalDatarootpath,cfgfile.主机名号 )
# # print( len(GenDict ) )


OrgfieldNames="归一牌号	归一牌号	IP	通道号(归一牌号[:10])	子防区号(顺序编号)	拟合光程	拟合光程	经度	纬度	经度	纬度	岸别	地名	相间距离 归属部门"
RegionFieldNames="定位界标1	定位界标2	电子围栏主机IP	通道号	子防区号	定位界标1光程	定位界标2光程	防区起点GPS经度坐标	防区起点GPS纬度坐标	防区终点GPS经度坐标	防区终点GPS纬度坐标	左右岸	地名信息	界标之间围栏长度 备注"

def getHostID(_GenDict):
	for key in _GenDict.keys():
		return key[:9]
	return None
import collections

def GenRegions( GenDict):
	# hostID = getHostID(GenDict )
	HostIP = getIPFrom主机名号( HostInfoDictArrary ,cfgfile.主机名号)
	hostID = getHostIDFrom主机名号(HostInfoDictArrary ,cfgfile.主机名号)
	hostID =hostID[1:]
	if hostID == None:
		print( "标牌编号存在错误，请检查后再处理")
		return None
	防区规划sdict=collections.OrderedDict()# 
	SigCnt= len( GenDict )	
	for chid in range(1,3,1):
		# RegionID =1
		for sigID in range(1,1000,1):
			Sig1_IDstr =hostID+str(chid )+ "_{:0>3}".format( sigID)
			Sig2_IDstr =hostID+str(chid )+ "_{:0>3}".format( sigID+1)
			Sig1_Dict =GenDict[ Sig1_IDstr ] 
			Sig2_Dict =GenDict[ Sig2_IDstr ] 
			if Sig1_Dict == None or Sig2_Dict == None:
				print( "处理通道{:}时在标牌{:}--{:}处结束。".format(chid,Sig1_IDstr,Sig2_IDstr ) )
				break
			regionDict = 光程DictCls()
			regionDict["定位界标1"] = "_"+Sig1_Dict[ "归一牌号"].replace("_","")
			regionDict["定位界标2"] = "_"+Sig2_Dict[ "归一牌号"].replace("_","")
			regionDict["电子围栏主机IP"] = HostIP
			regionDict["通道号"] = "_"+hostID+str(chid )
			regionDict["子防区号"] = str( sigID )
			regionDict["定位界标1光程"] = "{}".format(Sig1_Dict[ "拟合光程"] ); #print(int( Sig1_Dict[ "拟合光程"]  ) )
			regionDict["定位界标2光程"] = "{}".format(Sig2_Dict[ "拟合光程"] )
			regionDict["防区起点GPS经度坐标"] = Sig1_Dict[ "核准经度"]
			regionDict["防区起点GPS纬度坐标"] = Sig1_Dict[ "核准纬度"]			
			regionDict["防区终点GPS经度坐标"] = Sig1_Dict[ "核准经度"]
			regionDict["防区终点GPS纬度坐标"] = Sig1_Dict[ "核准纬度"]	
			regionDict["左右岸"] = Sig1_Dict[ "岸别"]
			location =Sig1_Dict.get("地名",""	)
			if "修正坐标" in location:
				regionDict["地名信息"] = ""
			else:
				regionDict["地名信息"] = location				
			# regionDict["地名信息"] = Sig1_Dict.get( "地名","")
			regionDict["界标之间围栏长度"] = Sig1_Dict[ "相间距离"]	


			regionDict["备注"] = Sig1_Dict.get("归属部门",""	)
			防区规划sdict [ Sig1_IDstr ]  =  regionDict  
		if Sig2_Dict == None:
			location =Sig1_Dict.get("地名",""	)
			if "修正坐标" in location:
				regionDict["地名信息"] = ""
			else:
				regionDict["地名信息"] = location
			# regionDict["地名信息"] =Sig1_Dict.get( "地名","")
	return 防区规划sdict;

ordered_d_fangqubiao= GenRegions(GenDict )
标牌数量=  len(GenDict ) 
防区数量= len( ordered_d_fangqubiao )
if 标牌数量-2 != 防区数量:
	print( "标牌数量和防区数量不符，请检查后处理！！！")
print( "标牌数量:",标牌数量)
print( "防区数量:",防区数量)

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

xlstgrName = fileProc.tplProc(cfgfile.root文件夹,"G002" ,cfgfile.主机名号,"防区规划","../基础数据\\reportTpl.xlsx")
Odbc2ADLib.xlstblWrite.saveDD2Xls(  xlstgrName,"防区规划",ordered_d_fangqubiao	,RegionFieldNames.split())
