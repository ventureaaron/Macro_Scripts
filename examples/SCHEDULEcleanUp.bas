Attribute VB_Name = "cleanUp"
Sub cleanSchedule()
'
' cleanSchedule Macro
'

'delete unwanted columns and row
'
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("K:K").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("R:R").Select
    Selection.Delete Shift:=xlToLeft
    Columns("R:R").Select
    Selection.Delete Shift:=xlToLeft
	
'remove values of 1 in session
'
    Range("G1").Select
    Do Until Selection.Value = ""
        If Selection.Value = "1" Then
            Selection.Value = ""
        End If
        Selection.Offset(1, 0).Select
    Loop

'remove periods in missing instructor
'
    Range("I1").Select
    Do Until Selection.Value = ""
        If Selection.Value = "." Then
            Selection.Value = ""
        End If
        Selection.Offset(1, 0).Select
    Loop
	
'set column width
'
	Columns("E:E").ColumnWidth = 23.43
    Columns("F:F").ColumnWidth = 6.57
    Columns("I:I").ColumnWidth = 10.14
    Columns("J:J").ColumnWidth = 6.71
    Columns("K:K").ColumnWidth = 8.29
    Columns("L:L").ColumnWidth = 9.57
	Columns("M:M").ColumnWidth = 12
    Columns("N:N").ColumnWidth = 5.86
    Columns("P:P").ColumnWidth = 8
    Columns("Q:Q").ColumnWidth = 28
	
'change column to align right
'
	Columns("M:M").Select
	With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
	
'change time display
'
	Columns("K:L").Select
    Selection.NumberFormat = "[$-409]h:mm AM/PM;@"
	
    Range("A1").Select
	
'remove duplicates
'
	ActiveSheet.Range("$A$1:$Q$1011").RemoveDuplicates Columns:=Array(1, 6, 7, 10), Header:=xlYes
            
End Sub

