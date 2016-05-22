Attribute VB_Name = "emailcleanUp"
Sub cleanEmail()
'
' cleanEmail Macro
'


'insert column to hold full name
'
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

'insert names into new column
'
    Range("B2").Select
    Do Until Selection.Value = ""
        Selection.Offset(-1, -1).Value = Selection.Value & " " & Selection.Offset(0, 1).Value
        Selection.Offset(1, 0).Select
    Loop

'grab emails from other columns in preferred email or mark the missing for deletion
'
    Range("A1").Select
    Do Until Selection.Value = ""
		If Selection.Offset(1,3).Value <> "" Then
			Selection.Value = Selection.Value & " <" & Selection.Offset(1,3).Value & ">"
		ElseIf Selection.Offset(1,4).Value <> "" Then
			Selection.Value = Selection.Value & " <" & Selection.Offset(1,4).Value & ">"
		ElseIf Selection.Offset(1,5).Value <> "" Then
			Selection.Value = Selection.Value & " <" & Selection.Offset(1,5).Value & ">"
		ElseIf Selection.Offset(1,6).Value <> "" Then
			Selection.Value = Selection.Value & " <" & Selection.Offset(1,6).Value & ">"
		Else: Selection.Value = "DELETE"
		End If
        Selection.Offset(1, 0).Select
    Loop
	
'clean up the commas in a name that make the upload script mess up (won't take commas or apostraphes around email line)
'
	Range("A1").Select
    Do Until Selection.Value = ""
        Selection.Value = Evaluate("=SUBSTITUTE("""& Selection.Value & ""","","","""")")
        Selection.Offset(1, 0).Select
    Loop
	
	
'remove rows with no email
'
	Range("A2").Select
    Do Until Selection.Value = ""
        If Selection.Value = "DELETE" Then
			Selection.Delete Shift:=xlUp
			Selection.Offset(-1,0).Select
		End If
        Selection.Offset(1, 0).Select
    Loop

'remove columns that aren't used
'
	Columns("B:G").Select
	Selection.Delete Shift:=xlToLeft
	
'Done
'
    Range("A1").Select
End Sub

'Selection.Offset(-1, -1).Value = Evaluate("=CONCATENATE(""" & Selection.Value & ""","" "",""" & Selection.Offset(0, 1).Value & """)")