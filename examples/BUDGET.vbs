

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
	arg1 = "\categorized.csv"
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
		if instr(objFile.Name, "Export") then
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
	  Set CM = oVBC.Import(sCurPath & "\BUDGET.bas") 
      xl.Application.run "budget_categorize"
      xl.DisplayAlerts = False        
      xlBook.saved = True
	  xl.ActiveWorkbook.SaveAs sCurPath & arg1
      xl.activewindow.close
      xl.Quit

      Set xlBook = Nothing
      Set xl = Nothing

	End Sub 