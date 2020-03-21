#-*- coding: UTF-8
from __future__ import print_function
import math
from commonFuncs import *
import csv
import os
import codecs
import kmltpl

def getPaiChIDStr (  guiyiPaihao ):
	return guiyiPaihao[0:10 ]
def getPaiIDNum (  guiyiPaihao ):
	return int( guiyiPaihao[ -3: ] )
def genPaihao (ChIDstr,pIDNum ):
	return ChIDstr+ "_{:0>3}".format( pIDNum)
def genGuiyipaiID(HostID,chidNum,pIDNum):
	return str(HostID)+ "{:0>1}".format( chidNum)+ "_{:0>3}".format( pIDNum)
class jiebiaoGPSCls:
	def __init__(self,jiebiaoRows,HostID):
		self.jiebiaoRows= jiebiaoRows
		self.KMlhead=kmltpl.g_KMlhead
		self.KMlTail=kmltpl.g_KMlTail
		self.kmlPosStrfmt=kmltpl.g_kmlPosStrfmt_12
		self.trackHead=kmltpl.g_trackHead
		self.trackTail=kmltpl.g_trackTail
		self.KMlhead_guiji=kmltpl.g_KMlhead_guiji
		self.KMlTailguiji=kmltpl.g_KMlTailguiji
		self.HostID = HostID
	def linerGps(self,head,tail):
			chidStr = getPaiChIDStr( head )
			row1 = self.jiebiaoRows[head]
			row2 =	self.jiebiaoRows[tail]	
			tail2 = getPaiIDNum ( tail )
			head1 = getPaiIDNum ( head )		
			cnt = tail2 - head1
			if cnt == 1:
				pass
			else:
				j2= str2Float(row2.get("核准经度","0.0") )
				j1 =str2Float(row1.get("核准经度","0.0") )
				deltaJingdu = ( j2-  j1)/ cnt 
				w2= str2Float(row2.get("核准纬度","0.0") )
				w1 =str2Float( row1.get("核准纬度","0.0") )
				deltaJingdu = ( j2-  j1)/ cnt 
				deltaWeidu = (w2-w1)/ cnt	
				for Pid  in range( head1+1,tail2,1):	
					guiyipaihao =genPaihao( chidStr,Pid )
					# print("---------------",guiyipaihao)
					row = self.jiebiaoRows[guiyipaihao]
					j = j1 + deltaJingdu*(Pid-head1)	
					w = w1 + deltaWeidu *(Pid-head1)
					# self.jiebiaoRows[guiyipaihao]	
					row["核准经度"],row["核准纬度"] =j ,w  
	def 界标坐标拟合计算(self):
		HostID= self.HostID
		MaxPID= 1000
		head = -1
		tail = -1
		for 通道号 in range(0,3,1):
			isJianGe =False
			print("Processing channel,begin :",HostID,通道号)
			for 牌号 in range( 1,MaxPID,1) :
				cur牌号 =genGuiyipaiID(HostID,通道号,牌号)
				Cur标牌信息Dict= self.jiebiaoRows[cur牌号] 
				# print( "---",cur牌号 )
				if None == Cur标牌信息Dict:
					print( cur牌号 )
					print("Processing channel,End :",HostID,通道号,"Max 牌号:",cur牌号)
					break
				坐标信息=self.getGPS( Cur标牌信息Dict )
				经度 = 坐标信息[0] 
				# print( "---",cur牌号,经度)
				if 经度!= None:
					if isJianGe == False :
							head = 牌号
							isJianGe = True
					else:
						tail = 牌号
						self.linerGps(genGuiyipaiID(HostID,通道号,head),genGuiyipaiID(HostID,通道号,tail))
						head =tail
	def getGPS(self,jiebiaoRow):
		return (jiebiaoRow.get("核准经度",None), jiebiaoRow.get("核准纬度",None) );
	
	def outPointTrack(self,filename ):
		f1 = codecs.open( filename,"wt",encoding="utf8")
		print(self.KMlhead ,file=f1 )
		Pais = sorted( self.jiebiaoRows.keys() )
		for key in Pais:
			row = self.jiebiaoRows[key]
			序号str =row["归一牌号"]
			if  序号str == None :
				break
			posstr = self.kmlPosStrfmt.format(row["归一牌号"],str2Float(row.get("相间距离","0.0") ),
				str2Float(row.get("核准经度","")),str2Float(row.get("核准纬度","") ),row["归一牌号"],
				row.get("岸别",""),
				row.get("地名","")  )
			print(posstr,file=f1)
		t1= ""
		t2= ""
		for key in Pais:
			row = self.jiebiaoRows[key]
		# for row in self.jiebiaoRows:
			序号str =  row["归一牌号"] 
			if  序号str == "" :
				break
			t2 =getPaiChIDStr(  row ["归一牌号"]  )
			if t1 !=  t2:
				if t1 != "" :
					print(self.trackTail ,file=f1 )
					print(self.trackHead ,file=f1 )
				else:
					print(self.trackHead ,file=f1 )
				t1=t2
			print( str2Float(row.get("核准经度","0.0") ),str2Float(row.get("核准纬度","0.0") ),sep=",",end=",0.0 \n",file=f1)

		print(self.trackTail ,file=f1 )
		print(self.KMlTail ,file=f1 )
		f1.close()

	def cal界标坐标(self,rootpath="",name = "G010-标牌GPS.kml"):
		KMZfilename =os.path.join(rootpath,name )
		self.界标坐标拟合计算()
		self.outPointTrack(KMZfilename )
		os.startfile(KMZfilename)
		
		
		
