function compileDLL_ipy($outname,$files){
    $respfile="resp111111111111111.txt"
    out-file -FilePath $respfile
    foreach(  $i in $files){
        out-file -FilePath $respfile -append -inputobject $i 
    }
    $name="ipyc2.7.exe"
    ###################
    $para=("/out:$outname","@resp111111111111111.txt") 
    Start-Process -FilePath $name -ArgumentList $para
}



$files1= "LibSRC/backupmail.py","LibSRC/CSVProc.py","LibSRC/Odbc2ADLib.py","LibSRC/Shell.py","LibSRC/XLSCombine.py","LibSRC/Exdecision.py"
$dllname="AppAssembly.dll"

compileDLL_ipy $dllname $files1 
