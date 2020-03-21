#-*- coding: UTF-8
#-*- coding: UTF-16
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

import os
import sys
# import cfgfile
import Odbc2ADLib 
import Exdecision

class sd_DictCls (dict):
	def __missing__(self,key) :
		return None

class cfgfile():
	@classmethod
	def getFiles(cls):
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


		# 主机信息文件夹 = 主机信息文件夹.strip( )
		print("主机信息文件名")
		TypeIDstr="A100"
		Newestfile,fullpaths,commonPAth= Exdecision.get最新file(TypeIDstr ,None ,  cls.主机信息文件夹)
		cls.主机信息文件名 = Newestfile
		cls.SignalDatarootpath = cls.root文件夹
		print("标牌GPS表 文件名")
		TypeIDstr="G001"
		Newestfile,fullpaths,commonPAth= Exdecision.get最新file(TypeIDstr ,cls.主机名号 ,rootPath = cls.root文件夹)
		cls.标牌GPS表 = Newestfile


		print( "标牌归属部门表 文件名")
		TypeIDstr="A200"
		Newestfile,fullpaths,commonPAth= Exdecision.get最新file(TypeIDstr ,cls.主机名号 , cls.root文件夹)
		cls.标牌归属部门表 = Newestfile

		cls.标牌节点信息表 = cls.节点顺序表 =cls.节点GUID表 =cls.节点替换表 = cls.标牌归属部门表

cfgfile.getFiles()
print()

标牌GPS表_DictArrary = Odbc2ADLib.xlstbl2AD.xls2AD([cfgfile.标牌GPS表], ["标牌GPS"]);print()	
节点替换表_DictArrary = Odbc2ADLib.xlstbl2AD.xls2AD([cfgfile.节点替换表], ["节点替换表"]);print()
标牌节点信息表_DictArrary = Odbc2ADLib.xlstbl2AD.xls2AD([cfgfile.标牌节点信息表], ["标牌节点表"]);print()
标牌归属部门表_DictArrary = Odbc2ADLib.xlstbl2AD.xls2AD([cfgfile.标牌归属部门表], ["归属部门表"]);print()
节点GUID表_DictArrary = Odbc2ADLib.xlstbl2AD.xls2AD([cfgfile.节点GUID表], ["节点GUID表"]);print()
节点顺序表_DictArrary = Odbc2ADLib.xlstbl2AD.xls2AD([cfgfile.节点顺序表], ["节点顺序表"]);print()
HostInfoDictArrary = Odbc2ADLib.xlstbl2AD.xls2AD([cfgfile.主机信息文件名], ["5电子围栏信息表（围栏标段填写）"],SpanLines=2);print()


def 归一牌号(row):
	keys = row.keys()
	if "归一牌号" in keys:
		return
	elif "通道号" in keys and "牌号" in keys  :
		通道号 = row["通道号"]
		牌号 = row["牌号"]
		归一牌号 = 通道号[ -10:]+"_"+ 牌号[-3:]
		# print(归一牌号,通道号 ,牌号 )
		row[ "归一牌号" ] = 归一牌号
	else:
		pass

map(归一牌号, 标牌节点信息表_DictArrary)
map(归一牌号, 标牌归属部门表_DictArrary)

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
			return IPStr[-9:]

HostIP = getIPFrom主机名号(HostInfoDictArrary, cfgfile.主机名号 )	
# KeyName: 合并不同Dict时，Keyname相同时，则进行字段的合并
# Arrarys,Fields ：Arrarys dict of dict 的数组，两个数组的个数应该相同。
#fields : for example:[ ("f5") ,("f4","f3") ,("f2" ,"f1")   ]
def CombineDictofDicts(KeyName, Arrarys,Fields ):
	tgrDict = sd_DictCls()
	datasNum = len(Arrarys )
	for i in range( 0,datasNum ):
		fields = Fields[i]
# 		print( fields )
		arraryDict = Arrarys[ i]
		for srcdic in arraryDict:
			key =srcdic[ KeyName  ]
			# print("====",KeyName,key )
			dataDict = tgrDict[key ]
			if dataDict ==None:
				dataDict= sd_DictCls()
			# print(key,i)
			dataDict[ KeyName ] = key
			for field in fields:
				# print( field )
				dataDict[ field ] = srcdic.get(field ,"")
				dataDict[ "IP" ] =  HostIP #getIPFromHostInfoTbl( HostInfoDictArrary,key )
			tgrDict[ key ] = dataDict
	return tgrDict

# 归一化标牌计算(标牌GPS表_DictArrary ,PostPro1=PostPro)	
# KeyName: 合并不同Dict时，Keyname相同时，则进行字段的合并
# Arrarys,Fields ：Arrarys dict of dict 的数组，两个数组的个数应该相同。
#fields : for example:[ ("f5") ,("f4","f3") ,("f2" ,"f1")   ]
def CombineDicts_坐标(KeyName, Arrarys,Fields ):
	tgrDict = sd_DictCls()
	datasNum = len(Arrarys )
	for i in range( 0,datasNum ):
		fields = Fields[i]
		arraryDict = Arrarys[ i]
		for srcdic in arraryDict:
			key =srcdic[ KeyName  ]
			dataDict = tgrDict[key ]
			if dataDict ==None:
				dataDict= sd_DictCls()
			# print(key,i)
			dataDict[ KeyName ] = key
			for field in fields:
				# print( field )
				dataDict[ field ] = srcdic[ field ]
				
			tgrDict[ key ] = dataDict
	return tgrDict

KeyName= "原节点ID"
Arrarys=[节点替换表_DictArrary ]
Fields=[("原节点ID","新节点ID","日期") ,  ]
节点替换表_DictofDict = CombineDicts_坐标(KeyName, Arrarys,Fields )


def ReplaceDictby节点ID( _节点替换表_DictofDict, _节点DictArrary,ReplaceFileNAme):
	for row in  _节点DictArrary:
		value = row[ ReplaceFileNAme ] 
		# print( value )
		tgrValue = _节点替换表_DictofDict[ value ]
		while tgrValue !=None:
			value = tgrValue[ "新节点ID" ]
			tgrValue = _节点替换表_DictofDict[ value ]

		row[ ReplaceFileNAme ] = value

ReplaceDictby节点ID( 节点替换表_DictofDict , 节点顺序表_DictArrary ,"一体化节点ID");print()
ReplaceDictby节点ID( 节点替换表_DictofDict , 标牌节点信息表_DictArrary,"一体化节点ID");print()
ReplaceDictby节点ID( 节点替换表_DictofDict , 节点GUID表_DictArrary ,"一体化节点ID");print()

KeyName= "一体化节点ID"
Arrarys=[ 节点GUID表_DictArrary ]
Fields=[("一体化节点ID","一体化节点GUID") ,  ]
节点GUID表_DictofDict = CombineDicts_坐标(KeyName, Arrarys,Fields )

def combine2ArraryofDict( arraryofDict1,keyName1 = "一体化节点ID",Allfields1=[],dictofDict2=None,Allfields2=[ "一体化节点GUID"]):
	newArraryofDict =[]
	for row in  arraryofDict1:
		newdict =sd_DictCls()
		KeyValue = row[ keyName1 ] 
		newdict[ keyName1 ] = KeyValue;
		for keyName   in Allfields1:
			newdict[ keyName ] = row[ keyName ];
		tgrRow = dictofDict2[ KeyValue ]
		if tgrRow ==None:
			print(" keyID:{:} can not found,Please Check".format(KeyValue ))
			continue
		for keyName  in Allfields2:
			newdict[ keyName ] = tgrRow[ keyName ];
		newArraryofDict += [newdict]
	return  newArraryofDict


guid对照表arraryofdict=combine2ArraryofDict( 节点顺序表_DictArrary,keyName1 = "一体化节点ID",Allfields1=[],dictofDict2=节点GUID表_DictofDict,Allfields2=[ "一体化节点GUID"])

def gen节点GUID顺序DictofDict( _guid对照表arraryofdict,keyName1= "一体化节点ID"):
	节点GUID顺序DictofDict = sd_DictCls()
	节点GUIDIdx次序DictofDict = sd_DictCls()
	cnt = len( _guid对照表arraryofdict )
	lastDict =_guid对照表arraryofdict[0]
	First节点Dict = lastDict
	for idx  in  range(1,cnt,1):
		CurDict = _guid对照表arraryofdict[idx]
		lastDict[ "Next"] =CurDict
		lastDict[ "CAN适配器主机IP"] = HostIP
		lastDict[ "一体化节点顺序号"] = str( idx)
		节点GUIDIdx次序DictofDict[ idx ] = lastDict
		keyValue = lastDict[keyName1 ]
		节点GUID顺序DictofDict[ keyValue ] = lastDict
		lastDict = CurDict
	lastDict[ "Next"] =None
	lastDict[ "CAN适配器主机IP"] = HostIP
	lastDict[ "一体化节点顺序号"] = str( idx)
	节点GUIDIdx次序DictofDict[ idx ] = lastDict
	keyValue = lastDict[keyName1 ]
	节点GUID顺序DictofDict[ keyValue ] = lastDict
	# lastDict = CurDict
	return ( 节点GUID顺序DictofDict,First节点Dict,节点GUIDIdx次序DictofDict );
节点GUID顺序DictofDict ,First节点GUIdDict,节点GUIDIdx次序DictofDict= gen节点GUID顺序DictofDict(guid对照表arraryofdict)

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
		# xlsTplName = os.path.join(rootPath,tplXlsFileName)
		xlsTplName = tplXlsFileName
		tgrFile=cls.getDTFileName(PreFix,A2Prefix+"_"+A3Prefix)
		xlstgrName =os.path.join( rootPath,tgrFile)
		srcfileObj =System.IO.FileInfo(xlsTplName )
		srcfileObj.CopyTo( xlstgrName )
		return xlstgrName

class 防区GenCls:
	def __init__( self ):
		self.LogicalPos_TableDictArrary=None
		self.Gps_TableDictArrary =None
		self.Host_TableDictArrary=None
		self.LogicalPosNameStr="拟合光程"
		self.KeyName = "归一牌号"
		self.Pos1Name = "定位界标1光程"
		self.Pos2Name = "定位界标2光程"	
		pass
	def getRegionFieldNames(self):
		RegionFieldNames = "定位界标1 定位界标2 电子围栏主机IP 子防区编号 起点一体化节点ID 终点一体化节点ID 防区起点GPS经度坐标 防区起点GPS纬度坐标 防区终点GPS经度坐标 防区终点GPS纬度坐标 防区位置 地名信息 界标之间围栏长度 备注"
		# RegionFieldNames="定位界标1 定位界标2 电子围栏主机IP 通道号 子防区号 {:} {:} 防区起点GPS经度坐标 防区起点GPS纬度坐标 防区终点GPS经度坐标	防区终点GPS纬度坐标	左右岸	地名信息 界标之间围栏长度 备注".format(self.Pos1Name,self.Pos2Name)
		return 	RegionFieldNames
	def gen标牌wholeInfoDictofDict(self):
		Arrarys=[ self.LogicalPos_TableDictArrary ,self.Gps_TableDictArrary, self.Host_TableDictArrary]
		self._fieldsArrary=[ (self.LogicalPosNameStr,),("核准经度","核准纬度","相间距离","岸别","地名",),("归属部门",)   ] #"GPS测试",
		Fields=self._fieldsArrary 
		self.标牌wholeInfoDictofDict = CombineDictofDicts(self.KeyName, Arrarys,Fields )
				
	def getHostID(self):
		return getHostIDFrom主机名号( HostInfoDictArrary,cfgfile.主机名号 )	

	def GenRegions( self ):
		hostID = self.getHostID(  )
		if hostID == None:
			print( "标牌编号存在错误，请检查后再处理")
			sys.exit()
			return None
		self.防区规划sdict=dict()# 
		SigCnt= len( self.标牌wholeInfoDictofDict  )	
		for chid in range(0,3,1):
			# RegionID =1
			for sigID in range(1,2000,1):
				Sig1_IDstr =hostID+str(chid )+ "_{:0>3}".format( sigID)
				Sig2_IDstr =hostID+str(chid )+ "_{:0>3}".format( sigID+1)
				# print(Sig1_IDstr)
				Sig1_Dict =self.标牌wholeInfoDictofDict [ Sig1_IDstr ] 
				Sig2_Dict =self.标牌wholeInfoDictofDict [ Sig2_IDstr ] 
				if Sig1_Dict == None or Sig2_Dict == None:
					print( "处理通道{:}时在标牌{:}--{:}处结束。".format(chid,Sig1_IDstr,Sig2_IDstr ) )
					break

				regionDict = sd_DictCls()
				regionDict["定位界标1"] = "_"+Sig1_Dict[ "归一牌号"].replace("_","")
				regionDict["定位界标2"] = "_"+Sig2_Dict[ "归一牌号"].replace("_","")
				regionDict["电子围栏主机IP"] = Sig2_Dict[ "IP"]
				regionDict["通道号"] = "_"+hostID+str(chid )
				regionDict["子防区编号"] = str( sigID )

				pos1= Sig1_Dict[ self.LogicalPosNameStr ]
				if sigID ==1:
					regionDict[self.Pos1Name] =   pos1 
				else:
					regionDict[self.Pos1Name] = self.getNextPos( pos1 )

				pos2= Sig2_Dict[ self.LogicalPosNameStr ]
				# print()
				# print(self.LogicalPosNameStr, pos2 )
				# print()
				regionDict[self.Pos2Name] = pos2 # self.getNextPos( pos2 )

				regionDict["防区起点GPS经度坐标"] = Sig1_Dict[ "核准经度"]
				regionDict["防区起点GPS纬度坐标"] = Sig1_Dict[ "核准纬度"]			
				regionDict["防区终点GPS经度坐标"] = Sig1_Dict[ "核准经度"]
				regionDict["防区终点GPS纬度坐标"] = Sig1_Dict[ "核准纬度"]	
				regionDict["防区位置"] = Sig1_Dict[ "岸别"]
				regionDict["地名信息"] = Sig1_Dict[ "地名"]	
				regionDict["界标之间围栏长度"] = Sig1_Dict[ "相间距离"]	
				regionDict["备注"] = Sig1_Dict[ "归属部门"]	
				self.防区规划sdict [ Sig1_IDstr ]  =  regionDict  
	def OutAllRegionDatas( self ):	
		fieldnamesStr = self.getRegionFieldNames()

		Odbc2ADLib.xlstblWrite.saveDD2Xls(  self.xlstgrName,"一体化防区规划",self.防区规划sdict,fieldnamesStr.split())

	def outGUIDDatas(self):
		fieldnamesStr = "CAN适配器主机IP 一体化节点ID 一体化节点GUID 一体化节点顺序号"
		# fieldnamesStr = self.getRegionFieldNames()
		Odbc2ADLib.xlstblWrite.saveDD2Xls(  self.xlstgrName,"GUID对应表",self.in节点GUIDIdx次序DictofDict ,fieldnamesStr.split())

	def gen防区规划( self ):
		self.gen标牌wholeInfoDictofDict()
		self.GenRegions(  )
		标牌归属部门数量=  len(self.Host_TableDictArrary ) 
		防区数量= len( self.防区规划sdict )
		LPTable数量= len( self.LogicalPos_TableDictArrary )
		GPSTable数量= len( self.Gps_TableDictArrary )
		maxCNT = max( 标牌归属部门数量,LPTable数量, GPSTable数量 )
		MinCnt =min( 标牌归属部门数量,LPTable数量, GPSTable数量 )
		if  maxCNT != MinCnt :
			print( "地理位置数据的数量{:}、标牌的逻辑坐标的数量{:}、归属部门表的数量{:}不符，请检查。".format(GPSTable数量,LPTable数量,标牌归属部门数量))
		if LPTable数量-2 != 防区数量:
			print( "标牌数量和防区数量不符，请检查后处理！！！")
		print( "标牌数量:",maxCNT)
		print( "防区数量:",防区数量)

		self.xlstgrName = fileProc.tplProc(cfgfile.SignalDatarootpath,"G200" ,cfgfile.主机名号,"一体化防区规划","../基础数据\\reportTpl.xlsx")
		self.OutAllRegionDatas( )
		self.outGUIDDatas()

	def getNextPos( CurPosStr ):
		return CurPosStr

class  一体化防区GenCls( 防区GenCls ):
	def getNextPos(self, CurPosStr ):
		# print( CurPosStr,end="---;" )
		try:
			curDict = self.in节点GUID顺序DictofDict[ CurPosStr ]
			nextDict= curDict["Next"]
			nextPos = nextDict[ "一体化节点ID"  ]
			# print( nextPos )
			if nextPos == None:
				# print( CurPosStr )
				return CurPosStr
			else:
				# print( nextPos )
				return nextPos
		except:
			# print( CurPosStr )
			return CurPosStr
					
a = 一体化防区GenCls()
a.LogicalPos_TableDictArrary= 标牌节点信息表_DictArrary
a.Gps_TableDictArrary = 标牌GPS表_DictArrary
a.Host_TableDictArrary= 标牌归属部门表_DictArrary
a.LogicalPosNameStr="一体化节点ID"
a.KeyName = "归一牌号"
a.Pos1Name = "起点一体化节点ID"
a.Pos2Name = "终点一体化节点ID"
a.主机IP信息表 = HostInfoDictArrary
a.in节点GUID顺序DictofDict =节点GUID顺序DictofDict
a.inFirst节点GUIdDict =First节点GUIdDict
a.in节点GUIDIdx次序DictofDict =节点GUIDIdx次序DictofDict
a.gen防区规划(   )
