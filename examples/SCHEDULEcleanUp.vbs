'for place i work, this file takes an incoming excel file from our SIS and produces a ready to publish on our website outgoing schedule file.
'We've since moved to a different procedure but at the time this would help our registrar with something he had to do several times a day.

'this script is to clean up delivered course schedule file. Should have this script scheduled, and as is should reside in same location as delivered file and as cleanUp.bas module.
'This is currently setup to interact all within the same directory folder: input file, output file, this script, and cleanUp.bas. Also, some security on Excel may need to be adjusted
'to allow VB module/program access.

Option Explicit
	
	'Parameter 1 to script: name of output file, if no param is passed, it is set to arg1 value, should be in form of "\filename.ext"
	
	'setup output file name and current path
	Dim args
	Dim arg1
	Dim fso
	Dim folo
	Dim folcolo
    Dim sCurPath
	Dim objFile
	
	Set args = WScript.Arguments
	arg1 = "\Spring_Schedule_2016.xls"
	if args.Count>0 then
		arg1 = args.Item(0)
	end if
	Set fso = CreateObject("Scripting.FileSystemObject")
	sCurPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")
	Set folo = fso.GetFolder(sCurPath&"\")
	Set folcolo = folo.Files
	
	'LaunchMacro to clean up feed file and save to output.
	'Changes to actual adjustment of file cleanup should be made at the cleanUp.bas module file, under the cleanSchedule macro.
	For Each objFile in folcolo
		if instr(objFile.Name, "UTZ2_P_SR_SCHEDULE_") then
		'cleanup for new output
			If fso.FileExists(sCurPath & arg1) then
				fso.DeleteFile sCurPath & arg1,true
			End If
			LaunchMacro
		'cleaunup for next feed file
			fso.DeleteFile objFile,true
		end if
	Next

    Sub LaunchMacro() 
      Dim xl
      Dim xlBook      
      Dim oVBC
      Dim CM

      Set xl = CreateObject("Excel.application")
      Set xlBook = xl.Workbooks.Open(sCurPath & "\" & objFile.Name, 0, True)      
      xl.Application.Visible = True
	  Set oVBC = xlBook.VBProject.VBComponents 
	  Set CM = oVBC.Import(sCurPath & "\SCHEDULEcleanUp.bas") 
      xl.Application.run "cleanSchedule"
      xl.DisplayAlerts = False        
      xlBook.saved = True
	  xl.ActiveWorkbook.SaveAs sCurPath & arg1
      xl.activewindow.close
      xl.Quit

      Set xlBook = Nothing
      Set xl = Nothing

	End Sub 