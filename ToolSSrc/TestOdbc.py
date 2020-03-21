# -*- coding: UTF-8
from __future__ import print_function
import clr
import System
import System.IO

import sys
clr.AddReference("System.Data")
import System.Data.Odbc


class TestCls():
	def get(self):
		self.xlsfilename= r"C:\nsbd\ipy\电子围栏Tools\tools\定时调测数据\xxxxx_2017年01月03日15_44_47.xlsx"
		self.RowBegin=0

		pattern2='''Driver={:};Provider=Microsoft.ACE.OLEDB.12.0;DBQ={:};IMEX=0;IgnoreCalcError=true;AllowFormula=false;Extended Properties="Excel 12.0 Xml;HDR=YES;EmptyTextMode=NullAsEmpty;IMEX=0";'''
		
		pattern32='''Driver={:};DBQ={:};IgnoreCalcError=true;AllowFormula=false;Extended Properties="Excel 12.0 Xml;HDR=YES;EmptyTextMode=NullAsEmpty;IMEX=0";'''
		pattern32='''DSN=TestX86'''

		DriverName = '{Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}'
		connectionString2 = pattern2.format(DriverName,self.xlsfilename )
		connectionString32 = pattern32.format(self.xlsfilename )
		# connectionString= pattern12
		connectionString = connectionString2
		print(self.xlsfilename  )

		print( connectionString )
		self.connection = System.Data.Odbc.OdbcConnection(connectionString)
		# print(connection.Database, connection.DataSource )
		try:
			self.connection.Open()
		except Exception as e:
			print( e )
			print(connectionString)

			sys.exit()
		tblinfo = self.connection.GetSchema("Tables") 
		self.Tbls=[]
		for row in tblinfo.Rows:
			for col in tblinfo.Columns:
				# print("----")
				if  col.ColumnName == "TABLE_NAME":
					#print(col.ColumnName )
				#print("===")
					# print(row[col])
					tblname = row[col]
					if tblname[-1:] == "$" or tblname[-2:] == "$'":
						# if tblname
						# print("-----------",tblname)
						self.Tbls +=[tblname]
						# print( tblname)
a=TestCls()
a.get()