' script file to print the results RTF file to the default printer
' DO NOT EDIT THIS FILE 
'
'Instead make a copy of it, rename it, then edit that - and then
' select the modified file in the OpenDMIS dialog

' get the full path and name of the RTF file
res = CamioUtils.GetResultsFilename(res_file)


'Create the WScript shell object
Set objWScript = CreateObject("WScript.Shell")


' file to print with quotes around it
 param1 = """" + res_file + """"
'param1 = res
 

' command line (with a hardcoded path to the wordpad executable)
strLaunchCmd = """D:\LGE_Report\Program\PDFCreator.exe"" " + param1

'Run the command
Set objExec = objWScript.Exec(strLaunchCmd)
'MsgBox "finish"

