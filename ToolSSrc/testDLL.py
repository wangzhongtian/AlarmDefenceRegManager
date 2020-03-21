#—-* coding:UTF-8
import clr 
import sys
dllname="AppAssembly.dll"
# sys.path.append(".")
# sys.path.append("/media/wang/USBdata1t/data/DOCs/ipy/防区处理工具/ToolSSrc/LibSRC")
for i in sys.path: print i 
print dllname
clr.AddReferenceToFileAndPath( dllname )
# import Odbc2ADLib
import Exdecision
Exdecision.testDLL()