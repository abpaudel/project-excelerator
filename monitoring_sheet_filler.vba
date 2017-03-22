Private Sub CommandButton1_Click()

    Dim x As Workbook
    Dim y As Workbook
    
    Dim lastrow As Long
    
    Dim datefinal As Long
    Dim dateinitial As Long
    
    Dim file As String
    
    Dim started As String
    Dim name As String
    Dim age As String
    Dim gender As String
    Dim district As String
    Dim caste As String
    Dim mobown As String
    Dim pconscent As String
    
    Dim datecell As Range

    Dim dateinitial_st As String
    Dim datefinal_st As String
    
    Dim before As String
    
    Set y = Application.ActiveWorkbook

    file = Application.ActiveWorkbook.Path & "/" & "data.xlsx"
    
    dateinitial_st = "2/14/2017"
    datefinal_st = InputBox("Enter date as mm/dd/yyyy")
    dateinitial = DateValue(dateinitial_st)
    datefinal = DateValue(datefinal_st)
    
    before = Format(datefinal + 1, "mm/dd/yyyy")
    
    With y.Sheets(1).Range("4:4")
            Set datecell = .Find(What:=DateValue(datefinal_st), LookIn:=xlFormulas, LookAt:=xlWhole)
            If datecell Is Nothing Then
                MsgBox "Date not found"
            End If
    End With
    
    Set x = Workbooks.Open(file)
    
    lastrow = x.Sheets(1).Range("A1").CurrentRegion.Rows.Count
    
    x.Sheets(1).Range("$A$1:$AD$" & CStr(lastrow)).AutoFilter Field:=5, Criteria1:="<" & before, Operator:=xlAnd
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
    
    datecell.Offset(3, 0).Value = started
    datecell.Offset(4, 0).Value = name
    datecell.Offset(5, 0).Value = age
    datecell.Offset(6, 0).Value = gender
    datecell.Offset(7, 0).Value = district
    datecell.Offset(8, 0).Value = caste
    datecell.Offset(9, 0).Value = mobown
    datecell.Offset(10, 0).Value = pconscent
    datecell.Offset(11, 0).Value = pconscent

    x.Close savechanges:=False

End Sub


