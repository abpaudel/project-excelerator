'Copyright (c) Abhishek Paudel | apaudel.com.np

Private Sub CommandButton1_Click()

Dim x As Workbook
Dim y As Workbook

Dim i As Integer
Dim j As Integer

Dim file As String
Dim temp As String

Dim dist As String
Dim vdc As String

dist = Application.InputBox("Enter name of the district")

temp = Application.ActiveWorkbook.path & "/" & "List Template.xlsx"

file = Application.GetOpenFilename _
(Title:="Please choose a source file", _
FileFilter:="Excel Files *.xls* (*.xls*),")
Set x = Workbooks.Open(file)

x.Sheets(1).Select

x.Sheets(1).Cells.Select
With Selection
    .Font.Size = 14
    .VerticalAlignment = xlCenter
End With
x.Sheets(1).Columns("G:G").Select
With Selection
    .HorizontalAlignment = xlCenter
End With
    
Dim lastrow As Integer

lastrow = x.Sheets(1).Range("A1").CurrentRegion.Rows.Count

For i = 1 To lastrow
    For j = 2 To 9
        If x.Sheets(1).Cells(i, j).Font.Name <> "Preeti" Then x.Sheets(1).Cells(i, j).Font.Name = "Times New Roman"
    Next j
Next i

x.Sheets(1).Rows(1).Delete

Set y = Workbooks.Open(temp)

y.Sheets(1).Range("C3").Value = dist

j = 1

Do While x.Sheets(1).Cells(1, 1) <> ""
    
    vdc = x.Sheets(1).Cells(1, 1).Value
    
    y.Sheets(j).Copy After:=y.Sheets(y.Sheets.Count)
    y.Sheets(j).Name = vdc
    y.Sheets(vdc).Range("C4").Value = vdc
    
    With x.Sheets(1).UsedRange
    i = Application.CountIf(x.Sheets(1).Columns("A:A"), vdc)
    End With
    
    x.Sheets(1).Range(x.Sheets(1).Cells(1, 2), x.Sheets(1).Cells(i, 10)).Copy
    y.Sheets(vdc).Range(y.Sheets(vdc).Cells(9, 2), y.Sheets(vdc).Cells(i + 8, 10)).PasteSpecial xlPasteValues
    x.Sheets(1).Range(x.Sheets(1).Cells(1, 2), x.Sheets(1).Cells(i, 10)).Copy
    y.Sheets(vdc).Range(y.Sheets(vdc).Cells(9, 2), y.Sheets(vdc).Cells(i + 8, 10)).PasteSpecial xlPasteFormats
    
    Dim rRng As Range
    Set rRng = y.Sheets(vdc).Range(y.Sheets(vdc).Cells(9, 1), y.Sheets(vdc).Cells(i + 8, 10))
    rRng.BorderAround xlContinuous
    rRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    rRng.Borders(xlInsideVertical).LineStyle = xlContinuous

    y.Sheets(vdc).Select
    y.Sheets(vdc).Range("A9").Select
    Selection.AutoFill Destination:=y.Sheets(vdc).Range(y.Sheets(vdc).Cells(9, 1), y.Sheets(vdc).Cells(i + 8, 1))
    
    y.Sheets(vdc).Rows("9:" & i + 8).Select
    Selection.RowHeight = 23
    
    x.Sheets(1).Rows("1:" & i).EntireRow.Delete
    
    
    j = j + 1
    
Loop

Application.DisplayAlerts = False
y.Sheets(j).Delete
Dim fname As String
fname = Application.ActiveWorkbook.path & "\" & dist & " List.xlsx"
y.SaveAs Filename:=fname, FileFormat:=xlOpenXMLWorkbook, ConflictResolution:=True
    
x.Close savechanges:=False
y.Close savechanges:=False
    
MsgBox ("Successfully copied to " & fname)


End Sub




Private Sub CommandButton2_Click()

Dim x As Workbook
Dim y As Workbook

Dim i As Integer
Dim j As Integer

Dim file As String
Dim temp As String

Dim dist As String
Dim vdc As String

dist = Application.InputBox("Enter name of the district")

temp = Application.ActiveWorkbook.path & "/" & "Registration Form Template.xlsx"

file = Application.GetOpenFilename _
(Title:="Please choose a source file", _
FileFilter:="Excel Files *.xls* (*.xls*),")
Set x = Workbooks.Open(file)

x.Sheets(1).Select

x.Sheets(1).Cells.Select
With Selection
    .Font.Size = 14
    .VerticalAlignment = xlCenter
End With
x.Sheets(1).Columns("G:G").Select
With Selection
    .HorizontalAlignment = xlCenter
End With


x.Sheets(1).Columns("C:F").EntireColumn.Delete
x.Sheets(1).Columns("E:E").Cut
x.Sheets(1).Columns("C:C").Select
Selection.Insert Shift:=xlToRight

Dim lastrow As Integer

lastrow = x.Sheets(1).Range("A1").CurrentRegion.Rows.Count

For i = 1 To lastrow
    For j = 2 To 4
        If x.Sheets(1).Cells(i, j).Font.Name <> "Preeti" Then x.Sheets(1).Cells(i, j).Font.Name = "Times New Roman"
    Next j
Next i

x.Sheets(1).Rows(1).Delete

Set y = Workbooks.Open(temp)


y.Sheets(1).Range("C3").Value = dist

j = 1

Do While x.Sheets(1).Cells(1, 1) <> ""
    
    vdc = x.Sheets(1).Cells(1, 1).Value
    
    y.Sheets(j).Copy After:=y.Sheets(y.Sheets.Count)
    y.Sheets(j).Name = vdc
    y.Sheets(vdc).Range("C4").Value = vdc
    
    
    With y.Sheets(vdc).PageSetup
        .LeftFooter = "District: " & dist
        .CenterFooter = "VDC: " & vdc
    End With

    
    With x.Sheets(1).UsedRange
        i = Application.CountIf(x.Sheets(1).Columns("A:A"), vdc)
    End With
    
    
    x.Sheets(1).Range(x.Sheets(1).Cells(1, 2), x.Sheets(1).Cells(i, 4)).Copy
    y.Sheets(vdc).Range(y.Sheets(vdc).Cells(9, 2), y.Sheets(vdc).Cells(i + 8, 4)).PasteSpecial xlPasteValues
    x.Sheets(1).Range(x.Sheets(1).Cells(1, 2), x.Sheets(1).Cells(i, 4)).Copy
    y.Sheets(vdc).Range(y.Sheets(vdc).Cells(9, 2), y.Sheets(vdc).Cells(i + 8, 4)).PasteSpecial xlPasteFormats
    
    Dim rRng As Range
    Set rRng = y.Sheets(vdc).Range(y.Sheets(vdc).Cells(9, 1), y.Sheets(vdc).Cells(i + 8, 13))
    rRng.BorderAround xlContinuous
    rRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    rRng.Borders(xlInsideVertical).LineStyle = xlContinuous
    
    y.Sheets(vdc).Select
    y.Sheets(vdc).Range("A9").Select
    Selection.AutoFill Destination:=y.Sheets(vdc).Range(y.Sheets(vdc).Cells(9, 1), y.Sheets(vdc).Cells(i + 8, 1))
    
    y.Sheets(vdc).Select
    y.Sheets(vdc).Range("F9").Select
    Selection.AutoFill Destination:=y.Sheets(vdc).Range(y.Sheets(vdc).Cells(9, 6), y.Sheets(vdc).Cells(i + 8, 6))
    
    y.Sheets(vdc).Select
    y.Sheets(vdc).Range("H9").Select
    Selection.AutoFill Destination:=y.Sheets(vdc).Range(y.Sheets(vdc).Cells(9, 8), y.Sheets(vdc).Cells(i + 8, 8))
    
    y.Sheets(vdc).Select
    y.Sheets(vdc).Range("K9").Select
    Selection.AutoFill Destination:=y.Sheets(vdc).Range(y.Sheets(vdc).Cells(9, 11), y.Sheets(vdc).Cells(i + 8, 11))
    

    y.Sheets(vdc).Range(y.Sheets(vdc).Cells(9, 1), y.Sheets(vdc).Cells(i + 8, 4)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    
    
    y.Sheets(vdc).Range(y.Sheets(vdc).Cells(9, 11), y.Sheets(vdc).Cells(i + 8, 13)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    
    y.Sheets(vdc).Rows("9:" & i + 8).Select
    Selection.RowHeight = 45
    
    y.Sheets(vdc).Rows(i + 8 & ":" & i + 8).Select
    Selection.AutoFill Destination:=y.Sheets(vdc).Rows(i + 8 & ":" & i + 28), Type:=xlFillDefault
    y.Sheets(vdc).Range("A" & i + 9 & ":D" & i + 28).Select
    Selection.ClearContents
    
    x.Sheets(1).Rows("1:" & i).EntireRow.Delete
    
j = j + 1
    
Loop

Application.DisplayAlerts = False
y.Sheets(j).Delete
Dim fname As String
fname = Application.ActiveWorkbook.path & "\" & dist & " Registration Form.xlsx"
y.SaveAs Filename:=fname, FileFormat:=xlOpenXMLWorkbook, ConflictResolution:=True
    
x.Close savechanges:=False
y.Close savechanges:=False
    
MsgBox ("Successfully copied to " & fname)

End Sub


