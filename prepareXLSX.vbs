Sub AzureRBACInit()
'
' AzureRBACInit Macro
' Prepare the RBAC Workbook from the latest CSV via PowerShell
'
'

With ActiveWorkbook
    lastRow = GetLastUsedRow
    .Sheets(1).Name = "Security Groups"
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Roles"
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Perms"
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Raw"
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Security By Roles"
    .Sheets(2).Range(("A2:B" & lastRow)).Select
    .Sheets("Roles").Paste
    .Sheets(1).Range(("C2:C" & lastRow)).Copy
    .Sheets("Perms").Paste
    .Sheets(1).UsedRange.Cells.Copy
    .Sheets("Raw").Paste
    .Sheets("Raw").Rows(1).Delete Shift:=xlShiftUp

    
    With .Sheets(1)
        .Activate
        .Columns("A:B").EntireColumn.AutoFit
        .Columns("C:C").Delete Shift:=xlToLeft
        .ListObjects.Add(xlSrcRange, Range("$A$1:$B$44"), , xlYes).Name = "Table1"
        .ListObjects("Table1").ShowAutoFilterDropDown = False
        .ListObjects("Table1").ShowTableStyleRowStripes = True
        .ListObjects("Table1").TableStyle = "TableStyleDark2"
        .Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        .Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        
        With .Range("A1")
            .FormulaR1C1 = "Microsoft Azure Resource Manager"
            .Font.Size = 18
            .Font.Bold = True
        End With
     
        With .Range("A2")
            .FormulaR1C1 = "Built-In Security Roles"
            .Font.Size = 12
            .Font.Bold = True
        End With
        
        ActiveWindow.DisplayGridlines = False
    End With
     
    With .Sheets("Perms")
        .Activate
        .Range(("A1:A" & lastRow)).TextToColumns Destination:=.Range("A1"), _
        DataType:=xlDelimited, ConsecutiveDelimiter:=True, Comma:=True, Space:=False
        lastCol = GetLastUsedColumn
        
        For i = 2 To lastCol
            outRow = outRow + lastRow + 1
            currCol = ColLetter(i)
            .Range(currCol & "1:" & currCol & lastRow).Cut Destination:=.Range("A" & outRow)
        Next i
        
        lastRow = .Range("A" & .Rows.Count).End(xlUp).Row + 1
        .Range("A1:A" & lastRow).RemoveDuplicates Columns:=1, Header:=xlNo
        
        .Cells.Replace What:="Microsoft.", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
        .Range("A1").End(xlDown).Offset(1, 0).Delete Shift:=xlShiftUp
        .Sort.SortFields.Clear
        .Columns("A:A").Sort Key1:=Range("A1"), Order1:=xlAscending
    End With
   
End With

End Sub

Function GetLastUsedColumn()
    lastCol = Cells.Find(What:="*", _
            After:=Range("A1"), _
            LookAt:=xlPart, _
            LookIn:=xlFormulas, _
            SearchOrder:=xlByColumns, _
            SearchDirection:=xlPrevious, _
            MatchCase:=False).Column
    GetLastUsedColumn = lastCol
End Function

Function GetLastUsedRow()
    lastRow = Cells.Find(What:="*", _
            After:=Range("A1"), _
            LookAt:=xlPart, _
            LookIn:=xlFormulas, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlPrevious, _
            MatchCase:=False).Row
    GetLastUsedRow = lastRow
End Function

Function ColLetter(Column) As String
    If Column < 1 Then Exit Function
    ColLetter = ColLetter(Int((Column - 1) / 26)) & Chr(((Column - 1) Mod 26) + Asc("A"))
End Function
