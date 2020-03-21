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

dllname=r"""STDLIBalls.DLL""" ;clr.AddReferenceToFileAndPath( getFullNames(dllname) )
dllname=r"""AppAssembly.dll""" ;clr.AddReferenceToFileAndPath( getFullNames(dllname) )

# import sys
# sys.exit()
import os
from datetime import datetime, date, time

import Exdecision
import Odbc2ADLib
###################################################################
# import CSVProc 
import  collections
class 光程DictCls (collections.OrderedDict):
	def __missing__(self,key) :
		return None

def genAllIDs(firstID,NextID ):
	if firstID == None:
		return [ NextID ]
	AllIds=[]

	Chid = firstID[0:2]
	pID1 = int( firstID[2:] )
	pID2 = int( NextID[2:] )
	for pID in range(pID1,pID2 +1):
		sID = '{:0>3}'.format(pID)
		AllIds += [Chid+ sID]
	return AllIds

def getHostUNIQIdFromHostName( _HostInfoDictArrary,HostNameSt ):
	for row in _HostInfoDictArrary:
		host = row["管理处名称"]
		# IP = row["IP地址"]
		# 主机友好名 = row["主机友好名"] 
		if host in HostNameSt:
			IPStr = row["电子围栏主机编号"]
			return "A"+IPStr
	return "---"


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
		print("标牌光程校准文件s")
		TypeIDstr="G004"
		Newestfile,fullpaths,commonPAth= Exdecision.get最新file(TypeIDstr ,cls.主机名号 , cls.root文件夹,exceptFolders)
		cls.updateFiles  =[ Newestfile ,] 
		cls.SignalDatarootpath  = commonPAth
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
		xlsTplName = tplXlsFileName
		tgrFile=cls.getDTFileName(PreFix,A2Prefix+"_"+A3Prefix)
		xlstgrName =os.path.join( rootPath,tgrFile)
		srcfileObj =System.IO.FileInfo(xlsTplName )
		srcfileObj.CopyTo( xlstgrName )
		return xlstgrName
def ProcEntry():
	print()
	print()
	cfgfile.getFiles()

	# map(writeDict,Host_TableDictArrary )
	HostInfoDictArrary   = Odbc2ADLib.xlstbl2AD.xls2AD([cfgfile.主机信息文件名], ["5电子围栏信息表（围栏标段填写）"],SpanLines = 2);print() #SpanLines = 2)
	# map(writeDict,HostInfoDictArrary )

	g主机编号= getHostUNIQIdFromHostName( HostInfoDictArrary,cfgfile.主机名号 )	

	校准_光程表 = Odbc2ADLib.xlstbl2AD.xls2AD(cfgfile.updateFiles, ["1-光程调试结果","光程调试结果"]);print()

	校准排序光程排序表=光程DictCls()
	for  row in 校准_光程表:
		Id= row["归一牌号"] ;
		time =row["触发开始时间"]
	
		Avg= float( row.get("光程均值" ,"0.0"))
		Avg = int(Avg)
		if Avg <= 1 :
			frStr  = "校准光程的光程均值为0，取消该校准点的数据：{:},首选均值={:}，次选均值为:{:}"
			# print( frStr.format(Id,row["首选均值"],row["次选均值"]) )
			continue
		#################
		光程均值 = Avg
		newrow = dict( {
				"拟合光程":光程均值,
				"触发开始时间":row["触发开始时间"],
				"光程均值":Avg,
				"首选均值":row.get("首选均值",""),
				"次选均值":row.get("次选均值",""),
				"首选邻值":row.get("首选邻值","") ,
				"次选邻值":row.get("次选邻值","") } )
		aId= 校准排序光程排序表.get(Id ,None) 
		if aId == None:
			校准排序光程排序表[ Id ] = newrow
		else:
			if aId[ "触发开始时间" ] < time:
				校准排序光程排序表[ Id ] = newrow

	Ids= sorted(校准排序光程排序表.keys() )
	cnt= len( Ids)
	o=0
	print(o,cnt)
	for o in range( 0,cnt):
	 	# print("+++++",Ids[o][0:1])
		if Ids[o][0:1] != "1":
			print( o)
			break
	if len( Ids[o:]) != 0:
		IDs = (Ids[0:o],Ids[o:]) 
	else:
		IDs = (Ids[0:o] ,) 
		
	# print(IDs )
	for Sids in IDs:
		cnt = len( Sids )
		if cnt <= 1:
			print( "Not enough sample to compute the  FLs.----- Error.")
			print(Sids)
			print()
			continue;
		sid1 =Sids[0]
		f1 = 1.0*校准排序光程排序表[sid1]["拟合光程" ]
		f2=0.0
		for sid2 in Sids[1:]:
			allIDs= genAllIDs( sid1,sid2)
			allIDs=allIDs[1:-1]
			f2 = 校准排序光程排序表[sid2]["拟合光程" ]
			l = len( allIDs )	
			if l > 0:
				deltaf = (f2*1.0-f1)/(l+1.0)
			idx = 1.0
			for sid3 in allIDs :
				f3 = int( f1+ idx * deltaf)
				newrow = dict( {
				"拟合光程":f3}
				)
				校准排序光程排序表[ sid3 ] = newrow
				idx += 1.0
			f1= f2
			sid1= sid2

	光程拟合表=光程DictCls()	
	Ids= sorted(校准排序光程排序表.keys() )
 	for sID in Ids:
		RowDict = 校准排序光程排序表[ sID ]
		牌号= g主机编号[1:] + sID
		RowDict[ "归一牌号" ] = 牌号[1:]
		光程拟合表[ sID ]= RowDict

	Exdecision.is_exed_data(RIDMax= 100,gongzuoRate =10)
	gfieldNames = "归一牌号  拟合光程  原光程  差异  触发开始时间  光程均值 首选均值 次选均值 首选邻值 次选邻值" 

	xlstgrName = fileProc.tplProc(cfgfile.root文件夹,"G003" ,cfgfile.主机名号,"界标光程表","../基础数据\\reportTpl.xlsx")
	lastPos =0
	for k,v in 光程拟合表.items():
		if "001" in k:
			lastPos = v["拟合光程"]
			continue
		print(k, lastPos,v["拟合光程"])			
		lastPos = v["拟合光程"]
		# for k,v in v.items():
		# 	print(k,v)

	# Odbc2ADLib.xlstblWrite.saveDD2Xls(  xlstgrName,"界标光程表", 光程拟合表,gfieldNames.split())


ProcEntry()