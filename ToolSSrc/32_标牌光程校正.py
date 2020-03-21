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

import os
import datetime
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

		print()
		print("标牌光程校准文件s")
		TypeIDstr="G004"
		Newestfile,fullpaths,commonPAth= Exdecision.get最新file(TypeIDstr ,cls.主机名号 , cls.root文件夹,exceptFolders)
		cls.updateFiles  =[ Newestfile ,] 

		print( "标牌光程文件名")
		TypeIDstr="M003"
		Newestfile,fullpaths,commonPAth= Exdecision.get最新file(TypeIDstr ,cls.主机名号 ,  cls.root文件夹,exceptFolders)
		cls.标牌光程文件名  = Newestfile
		cls.SignalDatarootpath  = commonPAth
		cls.Signalfilename  = cls.标牌光程文件名


###################################################################

class 光程DictCls (dict):
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


def linefit(x , y):
	import math
	try:
		N = float(len(x))
		sx,sy,sxx,syy,sxy=0,0,0,0,0
		for i in range(0,int(N)):
			sx  += x[i]
			sy  += y[i]
			sxx += x[i]*x[i]
			syy += y[i]*y[i]
			sxy += x[i]*y[i]
		a = (sy*sx/N -sxy)/( sx*sx/N -sxx)
		b = (sy - a*sx)/N
		r = abs(sy*sx/N-sxy)/math.sqrt((sxx-sx*sx/N)*(syy-sy*sy/N))
		return a,b,r
	except:
		return 0,0,0	
def getFL_标牌( _标牌光程排序表 ,AllIds ):
	FLs=[]
	for ID in AllIds:
		# print("+++++",ID )
		fl1 = _标牌光程排序表[ ID ] 
		if fl1 == None:
			fl1 =0.0 
		FLs +=[ fl1 ]
	return FLs

def getFL_校准( _校准排序光程排序表 ,AllIds ):
	FLs=[]
	for ID in AllIds:
		# print(ID,",,,,,")
		if ID == None:
			 continue
		FLs +=[ int( _校准排序光程排序表[ ID ]["光程均值"] ) ]
	return FLs

def CalNew拟合光程(a,b,oldFLs ):
	newFLs =[]
	for f0 in oldFLs:
		# print("---", f0)
		newFLs += [ int( f0*a+b ) ]
	return newFLs
def Proc_1_Section(firstID,NextID,_校准排序光程排序表 ,_标牌光程排序表):
	# return 
	if firstID != None:
# 		print( firstID,NextID ,end=":")
		AllIds = genAllIDs(firstID,NextID )
# 		print( AllIds )
		oldFLs = getFL_标牌( _标牌光程排序表 ,AllIds )
		
		old_NiheFL = getFL_标牌( _标牌光程排序表 ,[firstID,NextID] )
		new_NiheFL = getFL_校准( _校准排序光程排序表 ,[firstID,NextID] )
		a,b,r= linefit(old_NiheFL,new_NiheFL)
# 		print()
# 		print(a,b,r )
# 		print(old_NiheFL,new_NiheFL,oldFLs  )
		newFLs = CalNew拟合光程(a,b,oldFLs )
# 		print("------")
# 		print( oldFLs )
# 		print( newFLs)
# 		print("------")
		idx=0
		for  sID in AllIds:
			RowDict = _校准排序光程排序表[ sID ]
# 			oldFLs  newFLs
			isNew = False
			if RowDict == None:
				isNew =  True
				RowDict = 光程DictCls() 
			RowDict["原光程"] = oldFLs[ idx ]
			RowDict["拟合光程"] = newFLs[ idx ]
			RowDict["差异"] = abs( newFLs[idx ] -oldFLs[idx ] )
				
			if isNew :
				 _校准排序光程排序表[ sID ] = RowDict	
			idx += 1 
	else:
# 		print("Only 1 Row Data！！！！！！！！！！")
# 		print( firstID,NextID ,end=":")
		AllIds = genAllIDs(firstID,NextID )
 		# print( "----------",AllIds )
# 		oldFLs = getFL_标牌( _标牌光程排序表 ,AllIds )
		old_NiheFL = getFL_标牌( _标牌光程排序表 ,AllIds )
		new_NiheFL = getFL_校准( _校准排序光程排序表 ,AllIds )
		idx = 0
		for  sID in AllIds:
			RowDict = _校准排序光程排序表[ sID ]
# 			oldFLs  newFLs
			isNew = False
			if RowDict == None:
				isNew =  True
				RowDict = 光程DictCls() 
			RowDict["原光程"] = old_NiheFL[ idx ]
			RowDict["拟合光程"] = new_NiheFL[ idx ]
			RowDict["差异"] = abs( new_NiheFL[idx ] -old_NiheFL[idx ] )
				
			if isNew :
				_校准排序光程排序表[ sID ] = RowDict	
			idx += 1 	

def Combine2Dicts( _校准排序光程排序表 ,_标牌光程排序表 ):
	oldIds= sorted(_标牌光程排序表.keys() )
 	for soldID in oldIds:
		newRowDict = _校准排序光程排序表[ soldID ]
		oldFL = _标牌光程排序表[ soldID ]
		if newRowDict == None: # 在新校准排序标中 无 此牌号的数据
			newRowDict = 光程DictCls() 
			newRowDict[ "拟合光程" ] = oldFL
			newRowDict[ "差异" ] = ''
			newRowDict["原光程"] = oldFL
			_校准排序光程排序表[ soldID ] = newRowDict
		else:
			newFL = newRowDict[ "拟合光程" ]
			newRowDict["原光程"] = oldFL
			newRowDict[ "差异" ] = oldFL - newFL

print()
print()
cfgfile.getFiles()

标牌光程表 = Odbc2ADLib.xlstbl2AD.xls2AD([ cfgfile.Signalfilename], ["02-界标光程表","界标光程表"],SpanLines = 0,MaxRecords=3000);print()

标牌光程排序表=光程DictCls()
try:
	for row in 标牌光程表:
		paihao = row["归一牌号"] 
		break
except:
	paihao= None
	pass

if paihao == None:
	for row in 标牌光程表:
		chId= row["通道号"] ;
		g主机编号= chId[ :-1]
		break
	for row in 标牌光程表:
		chId= row["通道号"] ;
		SigId =row["牌号"]
		拟合光程 =row["拟合光程"]
		xuhao=row["序号"]
		
		ID= chId[-1:]+"_"+ SigId[1:]
		标牌光程排序表[ID ] = int( float( 拟合光程) )
if paihao != None:
	for row in 标牌光程表:
		chId= paihao;
		g主机编号= chId[ :9]
		break
	for row in 标牌光程表:
		# chId= row["通道号"] ;
		# SigId =row["牌号"]
		拟合光程 =row["拟合光程"]
		# xuhao=row["序号"]
		p=row["归一牌号"]  
		ID= p[9:]
		# print( ",,,",ID )
		标牌光程排序表[ID ] = int( float( 拟合光程) )

校准_光程表 = Odbc2ADLib.xlstbl2AD.xls2AD( cfgfile.updateFiles, ["1-光程调试结果","光程调试结果"],SpanLines = 0,MaxRecords=3000);print()

校准排序光程排序表=光程DictCls()
for  row in 校准_光程表:
	Id= row["归一牌号"] ;
	time =row["触发开始时间"]

	aId= 校准排序光程排序表[ Id ] 
	Avg= int( float( row[ "光程均值" ] ) )
	if Avg <= 1 :
		frStr  = "校准光程的光程均值为0，取消该校准点的数据：{:},首选均值={:}，次选均值为:{:}"
		# print( frStr.format(Id,row["首选均值"],row["次选均值"]) )
		continue

	OrgFl =标牌光程排序表[ Id ]
	if OrgFl !=None:
		# print(Id, OrgFl,row["首选均值"] ,row["次选均值"] )
		首选差异= abs( OrgFl- float( row["首选均值"] ) ) 
		均值差异= abs( OrgFl- float( row["次选均值"]  ))
		if 首选差异 <= 均值差异:
			光程均值 = float( row["首选均值"] )
		else:
			光程均值 = float( row["次选均值"] )
	else:
		OrgFl =0.0
		光程均值 = float( row["光程均值"] )
	#################
	光程均值 = float( row["光程均值"] )
	newrow = dict( {
			"拟合光程":光程均值,
			"原光程":OrgFl,
			"差异":光程均值-OrgFl,
			"触发开始时间":row["触发开始时间"],
			"光程均值":光程均值 ,
			"首选均值":row["首选均值"],
			"次选均值":row["次选均值"],
			"首选邻值":row["首选邻值"] ,
			"次选邻值":row["次选邻值"] } )
	if aId == None:
		校准排序光程排序表[ Id ] = newrow
	else:
		if aId[ "触发开始时间" ] < time:
			校准排序光程排序表[ Id ] = newrow

Ids= sorted(校准排序光程排序表.keys() )
cnt= len( Ids)
for o in range( 0,cnt):
 	# print("+++++",Ids[o][0:1])
	if Ids[o][0:1] != "1":
		break
if len( Ids[o:]) != 0:
	IDs = (Ids[0:o],Ids[o:]) 
else:
	IDs = (Ids[0:o] ,) 
	
print(IDs )
for Sids in IDs:
	cnt = len( Sids )
	if cnt ==0:
		continue
	idx=0
	firstID=None
	NextID =None

	while idx < cnt :
		if idx == 0:
			NextID =  Sids[ idx ] 
# 			  = firstID
			idx+=1
			continue
		else:
			firstID = NextID
			NextID  = Sids[ idx ] 
			Proc_1_Section(firstID,NextID ,校准排序光程排序表 ,标牌光程排序表)
			idx +=1
	if idx <= 1 :
		Proc_1_Section(firstID,NextID,校准排序光程排序表 ,标牌光程排序表 )
	print()
Exdecision.is_exed_data(RIDMax= 100,gongzuoRate =10)
gfieldNames = "归一牌号  拟合光程  原光程  差异  触发开始时间  光程均值 首选均值 次选均值 首选邻值 次选邻值" 
Combine2Dicts( 校准排序光程排序表 ,标牌光程排序表)

# def getDTFileName(Prefix,filenameTpl):
# 	tpl = "{:}_{:}_{:0>4}年{:0>2}月{:0>2}日{:0>2}_{:0>2}_{:0>2}.xlsx"
# 	curDt = datetime.now()
# 	filename = tpl.format(Prefix,filenameTpl,curDt.year,curDt.month,curDt.day,curDt.hour,curDt.minute,curDt.second )
# 	return filename

# tplXlsFileName ="reportTpl.xlsx"
# xlsTplName = os.path.join(cfgfile.root文件夹,tplXlsFileName)
# tgrFile=getDTFileName("G003",cfgfile.主机名号+"_光程递推")
# xlstgrName =os.path.join( cfgfile.root文件夹,tgrFile)
# srcfileObj =System.IO.FileInfo(xlsTplName )
# srcfileObj.CopyTo( xlstgrName )
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

xlstgrName = fileProc.tplProc(cfgfile.root文件夹,"G003" ,cfgfile.主机名号,"光程递推","../基础数据\\reportTpl.xlsx")
Ids= sorted(校准排序光程排序表.keys() )
for sID in Ids:
	RowDict = 校准排序光程排序表[ sID ]
	牌号= g主机编号[1:] + sID
	RowDict[ "归一牌号" ] = 牌号[1:]
	# RowDict[ "归一牌号" ] = g主机编号 + sID

Odbc2ADLib.xlstblWrite.AD2Xls(xlstgrName, "界标光程表",校准排序光程排序表.values()  )

