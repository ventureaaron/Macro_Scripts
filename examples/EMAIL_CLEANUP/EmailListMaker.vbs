Option Explicit
	
	'setup match pattern for files
	Dim myRegExp
	Set myRegExp = New RegExp
	myRegExp.IgnoreCase = True
	myRegExp.Global = True
	myRegExp.Pattern = "^(.*?).xlsx$"
	
	
	'setup output file name and current path

	Dim arg1
	Dim fso
	Dim folo
	Dim folcolo
    Dim sCurPath
	Dim objFile
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	sCurPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")
	Set folo = fso.GetFolder(sCurPath)
	Set folcolo = folo.Files
	
	'BuildFiles to clean up feed file and save to output.
	'Changes to actual adjustment of file cleanup should be made at the cleanUp.bas module file, under the users_modify macro.
	For Each objFile in folcolo
		if myRegExp.Test(objFile.Name)then
			arg1 ="\email_list.txt"
		'cleanup for new output
			If fso.FileExists(sCurPath & arg1 ) then
				fso.DeleteFile sCurPath & arg1,true
			End If
			BuildFiles
		'cleaunup for next feed file if uncommented
		'	fso.DeleteFile objFile,true
		end if
	Next
	
	
    Sub BuildFiles() 
	'builds _user files off of bb export files
      Dim xl
      Dim xlBook
	  Dim xlSheet
	  Dim row_count
      Dim oVBC
      Dim CM
	  
      Set xl = CreateObject("Excel.application")
      Set xlBook = xl.Workbooks.Open(sCurPath & "\" & objFile.Name, 0, True)      
      xl.Application.Visible = True
	  xlBook.Activate	  
	  Set oVBC = xlBook.VBProject.VBComponents 
	  Set CM = oVBC.Import(sCurPath & "\scripts\EmailcleanUp.bas") 
      xl.Application.run "cleanEmail"
	  xl.DisplayAlerts = False        
      xlBook.saved = True
	  
	  xl.ActiveWorkbook.SaveAs sCurPath & arg1, -4158
      xl.activewindow.close
      xl.Quit
	  

      Set xlBook = Nothing
      Set xl = Nothing

	End Sub 
	
	
	