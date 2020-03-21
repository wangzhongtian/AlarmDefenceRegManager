# -*- coding: UTF-8
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

dllname=r"""STDLIBalls.DLL""" ;clr.AddReferenceToFileAndPath( getFullNames(dllname) )
dllname=r"""AppAssembly.dll""" ;clr.AddReferenceToFileAndPath( getFullNames(dllname) )

import os
# import datetime
from datetime import datetime, date, time
# from cfgfile import *
import Exdecision
import XLSCombine
# import Exdecision
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

		cls.主机信息文件夹=cls.root文件夹 #r"C:\Cur\定时调测数据"

		cls.主机信息文件夹 ="../基础数据";# cls.主机信息文件夹.strip( )
		# cls.主机信息文件夹 =os.path.abspath(cls.主机信息文件夹+"/..")
		cls.srcAlmfilenames=[]
		cls.SignalDatarootpath = cls.root文件夹
		cls.srcTriggerfilenames =[]
		cls.rootpath = cls.root文件夹
		exceptFolders=["_排除文件夹"]

		for folder in exceptFolders:
			print( folder)

		print ("告警日志文件")
		try:
			TypeIDstr="A005"#"A005"
			Newestfile,fullpaths,commonPAth= Exdecision.get最新file(TypeIDstr ,cls.主机名号, cls.root文件夹,exceptFolders)
			# print( "============",fullpaths)
			if fullpaths != None:
				cls.srcAlmfilenames = cls.srcAlmfilenames +fullpaths
				cls.SignalDatarootpath = commonPAth
		except:
			pass
		try:
			print()
			print ("触发记录文件")
			TypeIDstr="A004"#"A004"
			Newestfile,fullpaths,commonPAth= Exdecision.get最新file(TypeIDstr ,cls.主机名号 , cls.root文件夹,exceptFolders)
			if fullpaths != None:
				cls.srcTriggerfilenames = cls.srcTriggerfilenames + fullpaths
			cls.rootpath = commonPAth	
		except:
			pass
	
###################################################################
def Deltadatas(ds1,ds2):
	deltaDs =[]
	for d1 in ds1 :
		if d1 not in ds2:
			deltaDs +=[ d1,]
	return deltaDs

def getMaxhappend范围_1(MaxRecords,datas ="""5460, 5470, 5478, 5486, 5486, 5498, 5500, 5530"""):
 	ds1 = datas.split( )
 	# print( ds1)
	ds=[]
	for s in ds1:
		ds +=[ float( s.replace(",","").replace(" ","")) ]
	ds =sorted(ds)
	maxpossibles=  getMaxhappendDatas( ds )

	dsDelta = Deltadatas(ds,maxpossibles)
	secondpossibles=  getMaxhappendDatas( dsDelta )

	dsDelta2 = Deltadatas(dsDelta,secondpossibles)
	threepossibles = getMaxhappendDatas( dsDelta2 )

	# print( "***************\r\n",ds1 )
	return [maxpossibles,secondpossibles,threepossibles]
def getMaxhappend范围(MaxRecords,datas ="""5460, 5470, 5478, 5486, 5486, 5498, 5500, 5530"""):
 	ds1 = datas.split( )
 	# print( ds1)
	ds=[]
	for s in ds1:
		ds +=[ float( s.replace(",","").replace(" ","")) ]

	dataSet =sorted(ds)
	maxpossibles =[]
	result =[]
	while( dataSet != [] ):
		dataSet = Deltadatas(dataSet,maxpossibles)
		maxpossibles=  getMaxhappendDatas( dataSet )
			# return 
		# print(len( maxpossibles ) ,MaxRecords )
		if len( maxpossibles ) >  MaxRecords:
			# print("-----",maxpossibles)
			continue
		else:
			result +=[ maxpossibles  ]
	return result

def getMaxhappendDatas(ds):	## input float  Arrary 	 
	if len( ds ) < 1:
 		 return (0,)
	cntArray =[ ] 
	numbercnt = len(ds)
	for i in range(0,numbercnt ):
	  cntArray += [ [1,]]

  	for k in range(0 , numbercnt):
  		curNumber = ds[ k ]
	  	for j in range( 0,numbercnt):
	  		if k != j:
		  		if ds[j] >curNumber-25.0 and ds[j] < curNumber+25.0:
		  			info = cntArray[k]
		  			cntArray[k] =[ info[0]+ 1,]+ info[1:] +[ds[j] ,]
	id_max=0
	for k in range( 0 , numbercnt ):
		cnt = cntArray[k] [0]
		if cnt > cntArray[ id_max ] [0]:
			id_max = k
	a= cntArray[ id_max ][ 1: ] +[ ds[ id_max],]
	return  a

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

def getAll范围Log(time1,time2,alarmdict,deltaMax=0.0,channelId='1',MaxRecords=1000000,光程修正=0.0):
	st = ""
#ID	Type	ChannelNumber	DefenceArea	Location	
#GPS	Message	Level	DateTime	KeepTime	State	LastUpdateDateTime
	for k in alarmdict:
		v= alarmdict[k]
		# print( v)
		chID = v["ChannelNumber"]
		# print( "------",chID,type(k),type(chID) ,type(time1) )
		inrange = k > (time1 -deltaMax)  and k < (time2 +deltaMax) 
		if( inrange and  channelId ==  chID ):#
			# print("....",channelId , v[1] ,v[0])
			data = float( v["Location"] ) + 光程修正
			st=st+" "+ str(data) 
	# if 光程修正 >0.0 :
	# 	print( st ,"*********",光程修正)
	# if 13012380.0- time2 <0.1:
	# 	print( st )
	# 	print()
	a = getMaxhappend范围(MaxRecords,st)
	# if 13012380.0- time2 <0.1:
	# 	print( sorted(st) )
	# 	print(a)
	# 	print()
	return a

##################################

def getpai归一号( row):
	#[ begTime,endTime,通道号,标牌号
	cnt =3
	d="000"
	p1 =row["标牌号"]
	l1= len(p1)
	s1=""
	if l1 <3:
		l2= l1-3
		s1= d[l2:]+p1
	else:
		s1= p1
	return row["通道号"][-1:]+"_"+s1 
	pass

##################################
def getAvgStr_1( datas):
	avg1 =0.0
	cnt1= len(datas)
	if  cnt1 != 0:
		avg1 =  sum( datas  ) / cnt1 
	avgstr=str(int(avg1) )
	str1=str(datas ) 
	return [avgstr,str1]
def getAvgStr( datas):
	avg1 =0.0
	cnt1= len(datas)
	if  cnt1 != 0:
		avg1 =  sum( datas  ) / cnt1 
	avgstr=str(int(avg1) )
	# str1=str(datas ) 
	return avgstr 


def writeDict(Row):
	# return 
	for k,v in Row.items():
		# print( v[-5:-4] )
		# if k in [ "核准经度","归一牌号" ]:
			print(k,v  )
	print()
	pass
def writetuple(Row):
	for k in Row:
		# print( v[-5:-4] )
		# if k in [ "核准经度","归一牌号" ]:
			print(k )

def alarmAd2Dd(ad告警数据):
	dd_alarmdict=  dict()
	for row in  ad告警数据:
#ID	Type	ChannelNumber	DefenceArea	Location	
#GPS	Message	Level	DateTime	KeepTime	State	LastUpdateDateTime
		Dt = row["DateTime"]
		Location = row["Location"]
		ChannelNumber = row["ChannelNumber"]
		row["ChannelNumber"]= str( int( float(ChannelNumber) ) )
		# dataField["通道号"]= str( int(float(通道号 ) ) )
		Id =row["ID"]

		# print("---",Dt)
		Dt = Dt.replace("  "," ").replace("年","/").replace("月","/").replace("日","")
		datestr,timestr = Dt.replace("  "," ").split(" ")
		# print(datestr,timestr)

		# Exdecision.is_exed_date(datestr.replace("/","  "),)
		Exdecision.is_exed_data(datestr.replace("/","  "),RIDMax= 100,gongzuoRate =10)
		l= getsecondslen( datestr,timestr)
		# print(datestr,timestr,l)
		row["归一时间"]= l
		dd_alarmdict[ l ] =row# (Dt,ChannelNumber, Location,Id )
	return dd_alarmdict
def 显示处理(row):
	# gFields="ID	Type	ChannelNumber	DefenceArea	Location GPS	Message	Level	_DateTime	KeepTime	State	LastUpdateDateTime".split()
	标牌号 = row["归一牌号"][-3:]
	row["标牌号"] = "A"+标牌号

	通道号 = row["通道号"][-1:]
	row["通道号"] = "A"+通道号
def 触发ad2Dd_范围( ad现场数据 ):
	mibiaodict=   dict()
	for dataField in ad现场数据 :
		标牌号 = dataField["标牌号"]
		通道号 = dataField["通道号"]
		print("Proc:",int(float(通道号) ),"_",int(float(标牌号)),dataField.get("日期","2016/01/01") ,dataField.get("起始敲击时间","Format Error") ,dataField.get("最后一次敲击结束时间","Format Error") ,"---")
		dataField["标牌号"]= str( int( float(标牌号) ) )
		dataField["通道号"]= str( int(float(通道号 ) ) )
		# print( "====",dataField["通道号"] ,dataField["标牌号"] )
		#0左右岸	1标牌号	2 起始敲击时间	3 最后一次敲击结束时间	4 敲击次数	5 通道号	6 光程记录	7 日期

		Dt1= dataField["日期"]
		date1 ,time1= Dt1.split()
		dataField["日期"] = date1
		date1= date1.replace("年","/").replace("月","/").replace("日","")

		Startdate,startTime = dataField["起始敲击时间"].split()
		# print(date1,startTime)
		begTime = getTime(date1,startTime)
		dataField["起始敲击时间"] = startTime


		Startdate,startTime = dataField["最后一次敲击结束时间"].split()
		endTime = getTime(date1, startTime )
		dataField["最后一次敲击结束时间"] = startTime

		# print("----",通道号,标牌号)
		时间调整str=  dataField.get( "主机时间减去工人手表时间差","0.0")

		if type(时间调整str) ==  str:
			时间调整 = float(时间调整str)
		elif type(时间调整str) ==  int  or type(时间调整str) == float :
			时间调整 = float(时间调整str)
		else:
			时间调整 =0.0
			print( "主机时间减去工人手表时间差" ,"格式填写错误，应为数值格式",type(时间调整str) )		
		# if abs(时间调整) > 0.0: print( "-----",时间调整)
		dataField["触发开始时间"] = begTime + 时间调整
		dataField["触发结束时间"] = endTime + 时间调整

		光称修正str =  dataField.get( "光程修正","0.0")
		# print("-----", 光称修正str)
		光称修正 = float( 光称修正str )
		dataField["光程修正"] = 光称修正

		kDate = (begTime,dataField["通道号"] )
		mibiaodict[ kDate ] = dataField # [ begTime,endTime,通道号,标牌号,光缆盘号,米标,备注 ]
	
	return mibiaodict
# import datetime
# def getDTFileName(Prefix,filenameTpl):
# 	tpl = "{:}_{:}_{:0>4}年{:0>2}月{:0>2}日{:0>2}_{:0>2}_{:0>2}.xlsx"
# 	curDt = datetime.now()
# 	filename = tpl.format(Prefix,filenameTpl,curDt.year,curDt.month,curDt.day,curDt.hour,curDt.minute,curDt.second )
# 	return filename
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
		# xlsTplName = os.path.join(rootPath,tplXlsFileName)
		xlsTplName=tplXlsFileName
		tgrFile=cls.getDTFileName(PreFix,A2Prefix+"_"+A3Prefix)
		xlstgrName =os.path.join( rootPath,tgrFile)
		srcfileObj =System.IO.FileInfo(xlsTplName )
		srcfileObj.CopyTo( xlstgrName )
		return xlstgrName


def 名称处理(row):
	# gFields="ID	Type	ChannelNumber	DefenceArea	Location GPS	Message	Level	_DateTime	KeepTime	State	LastUpdateDateTime".split()
	keyname=["DateTime", "Level","GUID"]
	for kn in keyname:
		val= row.get(kn,None)
		if val !=None:
			row["'{}'".format( kn )] =val
			del row[ kn ]

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
	ad告警数据 =  Odbc2ADLib.xlstbl2AD.xls2AD(cfgfile.srcAlmfilenames, ["Sheet1","sheet1",],SpanLines = 0,MaxRecords=-1);print()			
	print("-----------------------")
	# print( len(ad现场数据 ))
	# map( writeDict,ad现场数据)
	# print( len(ad告警数据 ))
	# map( writeDict,ad告警数据)

	dd_alarmData = alarmAd2Dd( ad告警数据 )
	# print("---------")
	print("-------------ad告警数据s shuju Ok ----------")
	dd_现场数据 = 触发ad2Dd_范围(ad现场数据 )
	print("-----------ad现场数据 shuju Ok ------------")
	for kDate,row in dd_现场数据.items():
		channelId = row["通道号"]
		qiaojiCnt1 = float( row["敲击次数"] )
		qiaojiCnt =1000
		修正值= float( row["光程修正"] )
		neghbores = getAll范围Log(row["触发开始时间"],row["触发结束时间"],dd_alarmData,deltaMax=cfgfile.deltaMax,channelId= channelId,MaxRecords=qiaojiCnt,光程修正=修正值)

		titles= ["首选均值"] + ["首选邻值"]+ ["次选均值"]+ ["次选邻值"]+["三选均值"] +["三选邻值"] +["四选均值"] +["四选邻值"] +["五选均值"] +["五选邻值"] 
		idx =0;
		neighboureLen =len( neghbores)
		# cntArrry=[]
		# for dataArrary in neighbores:
		# 	cntArrry += [ len( dataArrary )]


		loopcnt= min(neighboureLen,3 )
		while idx < loopcnt:
			# print( idx, loopcnt,neighboureLen )
			datas = neghbores[ idx]
			v1 = getAvgStr(datas )
			row[ titles[2*idx] ] =v1
			row[ titles[2*idx+1] ] = datas
			idx += 1
		row["光程均值"] = row.get("首选均值",0.0)

		p1 = getpai归一号( row )
		row["归一牌号"] = p1


# gfieldNames = "归一牌号 通道号 牌号 岸别 核准经度 核准纬度 地名 测试经度 测试纬度 计算经度 计算纬度"
	map( 名称处理 ,dd_alarmData.values() )
	map( 显示处理,dd_现场数据.values())
	# print("=====================")
	# map( writeDict,dd_alarmData.values() )

	xlstgrName = fileProc.tplProc(cfgfile.root文件夹,"G004" ,cfgfile.主机名号,"现场实测","../基础数据\\reportTpl.xlsx")
	Odbc2ADLib.xlstblWrite.AD2Xls(xlstgrName, "原始调测数据",dd_现场数据.values() )
	Odbc2ADLib.xlstblWrite.AD2Xls(xlstgrName, "报警数据",dd_alarmData.values()  )

	gFieldName="归一牌号	通道号	标牌号	起始敲击时间	最后一次敲击结束时间	触发结束时间	触发开始时间	光程均值	首选均值	次选均值	三选均值	首选邻值	次选邻值	三选邻值	敲击次数	日期	主机时间减去工人手表时间差"

	ad_data = 现场调测整理(dd_现场数据 )
	# map( writeDict,ad_data )
	Odbc2ADLib.xlstblWrite.AD2Xls(xlstgrName, "光程调试结果",ad_data ,gFieldName.split())
def entry定时同步数据处理():
	cfgfile.getFiles()
	print("decisinon Make.........................")
	# Exdecision.is_exed_data(  RIDMax= 100,gongzuoRate =100)
	print("decision Ok...........................")
	时间区间分析法()

entry定时同步数据处理()