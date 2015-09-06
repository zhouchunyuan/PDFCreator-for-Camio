Dim objShell
Set objShell = WScript.CreateObject( "WScript.Shell" )
param1 ="""D:\LGE_Report\LTG\main.res""" 
strLaunchCmd = """D:\LGE_Report\Program\PDFCreator.exe"" " + param1
objShell.Exec(strLaunchCmd )
Set objShell = Nothing



