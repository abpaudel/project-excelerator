Private Sub CommandButton1_Click()

    Dim x As Workbook
    Dim y As Workbook
    
    Dim lastrow As Long
    Dim i As Long
    Dim j As Integer
    
    Dim file As String
    
    Dim started As String
    Dim name As String
    Dim age As String
    Dim gender As String
    Dim district As String
    Dim caste As String
    Dim mobown As String
    Dim pconscent As String
    
    
    Dim mydate As String

    
    file = Application.GetOpenFilename _
    (Title:="Please choose a source file", _
    FileFilter:="Excel Files *.xls* (*.xls*),")
    
    'mydate = InputBox("Enter date as mm/dd/yyyy")
    mydate = "3/15/2017"
    Set x = Workbooks.Open(file)
    
    lastrow = x.Sheets(1).Range("A1").CurrentRegion.Rows.Count
    
    x.Sheets(1).Range("$A$1:$AD$" & CStr(lastrow)).AutoFilter Field:=5, Criteria1:="<" & mydate, Operator:=xlAnd
    started = WorksheetFunction.Subtotal(103, ActiveSheet.Range("$A$1:$AD$" & CStr(lastrow)).Columns(1)) - 1
    x.Sheets(1).Range("$A$1:$AD$" & CStr(lastrow)).AutoFilter Field:=10, Criteria1:="<>"
    name = WorksheetFunction.Subtotal(103, ActiveSheet.Range("$A$1:$AD$" & CStr(lastrow)).Columns(1)) - 1
    x.Sheets(1).Range("$A$1:$AD$" & CStr(lastrow)).AutoFilter Field:=13, Criteria1:="<>"
    age = WorksheetFunction.Subtotal(103, ActiveSheet.Range("$A$1:$AD$" & CStr(lastrow)).Columns(1)) - 1
    x.Sheets(1).Range("$A$1:$AD$" & CStr(lastrow)).AutoFilter Field:=16, Criteria1:="<>"
    gender = WorksheetFunction.Subtotal(103, ActiveSheet.Range("$A$1:$AD$" & CStr(lastrow)).Columns(1)) - 1
    x.Sheets(1).Range("$A$1:$AD$" & CStr(lastrow)).AutoFilter Field:=19, Criteria1:="<>"
    district = WorksheetFunction.Subtotal(103, ActiveSheet.Range("$A$1:$AD$" & CStr(lastrow)).Columns(1)) - 1
    x.Sheets(1).Range("$A$1:$AD$" & CStr(lastrow)).AutoFilter Field:=22, Criteria1:="<>"
    caste = WorksheetFunction.Subtotal(103, ActiveSheet.Range("$A$1:$AD$" & CStr(lastrow)).Columns(1)) - 1
    x.Sheets(1).Range("$A$1:$AD$" & CStr(lastrow)).AutoFilter Field:=25, Criteria1:="<>"
    mobown = WorksheetFunction.Subtotal(103, ActiveSheet.Range("$A$1:$AD$" & CStr(lastrow)).Columns(1)) - 1
    x.Sheets(1).Range("$A$1:$AD$" & CStr(lastrow)).AutoFilter Field:=28, Criteria1:="<>"
    pconscent = WorksheetFunction.Subtotal(103, ActiveSheet.Range("$A$1:$AD$" & CStr(lastrow)).Columns(1)) - 1
    
    MsgBox started & "; " & name & "; " & age & "; " & gender & "; " & district & "; " & caste & "; " & mobown & "; " & pconscent
    x.Close savechanges:=False

End Sub
