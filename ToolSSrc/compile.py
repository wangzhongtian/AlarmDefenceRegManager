#_* coding:UTF-8
def compileDLL_ipy(outname,files):
    compileOptions={
                    # "mainModule":"aaa"  ,
                    "assemblyFileVersion":"0.0.0.0",
                    "copyright":"By wangzht",
                    "productName":"For Potevio",
                    "productVersion":"0.0.0.1"
                    }
    # import System 
    import clr 
    # files = ["LibSRC/backupmail.py","LibSRC/CSVProc.py","LibSRC/Odbc2ADLib.py","LibSRC/Shell.py","LibSRC/XLSCombine.py","LibSRC/Exdecision.py"]
    clr.CompileModules( outname ,*files ,**compileOptions);

def compileDLL_ipyc(outname,files):
    import os
    cmdstr1="$name='ipyc.exe'; $para=('/out:{}','@resp111111111111111.txt')".format( outname )
    cmdstr1 +=";Start-Process -FilePath $name -ArgumentList $para"
    respfileObj=open("resp111111111111111.txt","wt" )
    for i in files:
        respfileObj.write(i+"\n")
    print cmdstr1
    respfileObj.close()
    processid = os.system( cmdstr1 )
    # os.wait( processid)

def compileDLL_ipyc1(outname,files):
    import os
    cmdstr1="'ipyc.exe'"
    cmdstr2=r"'ipyc.exe' '/out:{}' '@resp111111111111111.txt')".format( outname )
    # cmdstr1 +=";Start-Process -FilePath $name -ArgumentList $para"
    respfileObj=open("resp111111111111111.txt","wt" )
    for i in files:
        respfileObj.write(i+"\n")
    print cmdstr1,cmdstr2
    respfileObj.close()
    os.spawnvp( os.P_WAIT, cmdstr1,cmdstr2 )
 

files1= ["LibSRC/backupmail.py","LibSRC/CSVProc.py","LibSRC/Odbc2ADLib.py","LibSRC/Shell.py","LibSRC/XLSCombine.py","LibSRC/Exdecision.py"]
dllname="AppAssembly.dll"
compileDLL_ipy( outname=dllname,files=files1)
print "dsfdf"
# import System 

