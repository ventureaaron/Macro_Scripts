'This is a general VBS script to open an excel file, associate a macro (bas) file with it, and launch the macro. The general
'flow is 1: find file to work on, 2: call macro on file, 3: save file.

Option Explicit
	
	'Parameter 1 to script: name of output file, if no param is passed, it is set to arg1 value, should be in form of "\filename.ext"
	'this is useful if you want to schedule this script and want to call the outbound file different things at different times of day
	'no args means the file with be named "outbound.csv". The default search for the name of the inbound file is to look for the file
	'containing "input" in its name, see comment at line 34-35 for how to change this.
	
	'setup output file name and current path
	'script call variables
	Dim args
	Dim arg1
	'file system object, folder object, folder collection object
	Dim fso
	Dim folo
	Dim folcolo
	'variables to hold location and name of file
    Dim sCurPath
	Dim objFile
	
	Set args = WScript.Arguments
	arg1 = "\outbound.csv"
	if args.Count>0 then
		arg1 = args.Item(0)
	end if
	Set fso = CreateObject("Scripting.FileSystemObject")
	sCurPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")
	Set folo = fso.GetFolder(sCurPath&"\")
	Set folcolo = folo.Files
	
	'Changes to actual adjustment of excel file should be made at the yourMacro.bas module file.
	'searches for each file in folder if name contains "Input" so if your inbound file you want to
	'manipulate is "Input.xlsx" this would be fine, otherwise change the "Input" string on line 37
	For Each objFile in folcolo
		if instr(objFile.Name, "Input") then
		'cleanup for new output
			If fso.FileExists(sCurPath & arg1) then
				'this makes sure there is no already existing outbound file and there won't be any problems saving a new one
				fso.DeleteFile sCurPath & arg1,true
			End If
			'this sub routine opens the excel file, pulls the macro and runs it and saves the outbound file
			LaunchMacro
			'cleaunup for next feed file, if you don't want it to remove the inbound file, then comment the next line out
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
	  Set CM = oVBC.Import(sCurPath & "\yourMacro.bas") 
      xl.Application.run "your_macro_name"
      xl.DisplayAlerts = False        
      xlBook.saved = True
	  xl.ActiveWorkbook.SaveAs sCurPath & arg1
      xl.activewindow.close
      xl.Quit

      Set xlBook = Nothing
      Set xl = Nothing

	End Sub 