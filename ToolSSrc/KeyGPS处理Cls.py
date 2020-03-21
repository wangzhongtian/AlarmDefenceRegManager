#-*- coding: UTF-8
from __future__ import print_function
import math
from commonFuncs import *
import csv
import os
import kmltpl
# from  XLSCombine  import *
# from  CSVFileOPs import *
import csv, codecs, cStringIO
class keyGPSroccls:
	def __init__(self,GPSRows):
		self.GPSlst=[]
		self.KMlhead=kmltpl.g_KMlhead
		self.KMlTail=kmltpl.g_KMlTail
		self.kmlPosStrfmt=kmltpl.g_kmlPosStrfmt_3
		self.trackHead=kmltpl.g_trackHead
		self.trackTail=kmltpl.g_trackTail
		self.KMlhead_guiji=kmltpl.g_KMlhead_guiji
		self.KMlTailguiji=kmltpl.g_KMlTailguiji
		self.KeyGPSRows =GPSRows

##序号	通道号	牌号	岸别	地名	测试经度	测试纬度	计算经度	计算纬度	核准经度	核准纬度
	def outPointTrack(self,filename ="c:\\a\\防区规划坐标.kml"):
		f1 = codecs.open( filename,"wt",encoding="utf8")
		print(self.KMlhead ,file=f1 )
		for row in self.KeyGPSRows:
			# if len(row.get("地名","") )>0:
			
				posstr = self.kmlPosStrfmt.format(row["归一牌号"],str2Float(row["核准经度"]),str2Float(row["核准纬度"]), \
						"",row.get("岸别","") ,row.get("地名","")  )
				print(posstr,file=f1)
		print(self.trackHead ,file=f1 )
		for row in self.KeyGPSRows:
				print( str2Float(row["核准经度"]),str2Float(row["核准纬度"]),sep=",",end=",0.0 \n",file=f1)
		print(self.trackTail ,file=f1 )
		print(self.KMlTail ,file=f1 )
		f1.close()
	def get当岸坐标(self,currow ,岸别lst):
		unitdegree= 80.0/30.0/3600.0
		posname = currow.get("地名","_").strip()
		fuhao = -1.0
		if "左" in 岸别lst:
			fuhao =1.0
		for row in self.KeyGPSRows:
			curname = row.get("地名","_").strip()
			if row.get("岸别","_").strip().lower()  in 岸别lst:
				if posname ==  curname:
					jingdu = str2Float( row.get("测试经度","0.0") ) + fuhao*unitdegree 
					weidu =str2Float( row.get("测试纬度","0.0") ) + fuhao*unitdegree 
					return (jingdu,weidu)
		return 	 ("114.45","37.35")		
			
		
	def  get对岸坐标(self,currow):
		可能方向左=("left","左")
		可能方向右=("right","右")
		if currow.get("岸别","_").strip().lower() not in 可能方向左:
			return self.get当岸坐标(currow ,可能方向左)
		else:
			return self.get当岸坐标(currow ,可能方向右)
	
		
	def  计算对岸坐标(self):
		for row in self.KeyGPSRows:
			if row.get("核准经度","").strip()=="":
				if(row.get("测试经度","").strip() ) !="":
					row["核准经度"] = row.get("测试经度","").strip()
				else:
					row["核准经度"] = self.get对岸坐标(row)[0]
			if row.get("核准纬度","").strip() =="":
				if(row.get("测试纬度","").strip() ) !="":
					row["核准纬度"] = row.get("测试纬度","").strip()					
				else:
					row["核准纬度"] = self.get对岸坐标(row)[1]
	def show界标坐标(self,rootpath="",name="71-关键点坐标.kml"):
		self.计算对岸坐标()
		KMZfilename =os.path.join(rootpath,name)
		self.outPointTrack(KMZfilename )
		os.startfile(KMZfilename)
