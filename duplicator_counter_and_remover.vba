Private Sub CommandButton1_Click()

Dim x As Workbook
Dim fname As String
Dim dist As String

Dim i As Long
Dim c As Integer

file = Application.GetOpenFilename _
(Title:="Please choose a source file", _
FileFilter:="Excel Files *.xls* (*.xls*),")
Set x = Workbooks.Open(file)



dist = x.Sheets(1).Cells(2, 3).Value

i = 2

Do While x.Sheets(1).Cells(i, 2) <> ""
    c = 1
    Do While x.Sheets(1).Cells(i, 2) = x.Sheets(1).Cells(i + 1, 2)
        c = c + 1
        x.Sheets(1).Rows(i + 1).Delete
        
    Loop
    x.Sheets(1).Cells(i, 7).Value = c
    i = i + 1
Loop

Application.DisplayAlerts = False

fname = Application.ActiveWorkbook.Path & "\" & dist & ".xlsx"
x.SaveAs Filename:=fname, FileFormat:=xlOpenXMLWorkbook, ConflictResolution:=True
x.Close savechanges:=False

MsgBox ("Successfully copied to " & fname)
End Sub
