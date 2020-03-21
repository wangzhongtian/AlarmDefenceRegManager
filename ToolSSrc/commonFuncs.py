#-*- coding: UTF-8
import math
import os

def getFileFromSheetName(sheetname ,curpath):
	fileName = os.path.join(curpath,sheetname+"_OO.csv" )
	return fileName
def str2Float(StrData):
	try:
		return float( StrData )
	except:
		return 0.0
	
def linefit(x , y):
# 	print(x,len(x))
# 	print(y,len(y))
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

def isInRange(panhaoStr,Mibiao,panhaoStr1,mibiao1,panhaoStr2,mibiao2):
	pStr= panhaoStr.strip() 
	皮长=-1.0
	if pStr  == panhaoStr1.strip() and pStr  == panhaoStr2.strip() :
		if Mibiao >=  mibiao1 and Mibiao <=  mibiao2:
			皮长= Mibiao -  mibiao1
		if Mibiao <= mibiao1 and Mibiao  >=  mibiao2:
			皮长= mibiao1 -  Mibiao 		
			
# 	if (皮长>=0.0):
# 		print(panhaoStr,Mibiao,panhaoStr1,mibiao1,panhaoStr2,mibiao2,皮长)
	return 皮长
# jiezhiriqi="2016 8 10"
 
  