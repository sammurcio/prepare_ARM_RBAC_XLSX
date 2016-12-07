Sub AzureRBACInit()
'
' AzureRBACInit Macro
' Prepare the RBAC Workbook from the latest CSV via PowerShell
'
'

With ActiveWorkbook
    lastRow = GetLastUsedRow(ActiveSheet.Name)
    .Sheets(1).Name = "Security Groups"
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "temp_roles"
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "temp_perms"
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "temp_raw"
    .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "raw_table"
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
        lastCol = GetLastUsedColumn(.Name)
        
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
    
    With .Sheets("raw_table")
        Set roleSh = ActiveWorkbook.Sheets("Roles")
        Set permsh = ActiveWorkbook.Sheets("Perms")
        Set rawSh = ActiveWorkbook.Sheets("Raw")
        .Activate
        .Cells(1, 1).Value = "Role"
        .Cells(1, 2).Value = "Permissions"
        .Cells(1, 3).Value = "Enabled"
        permCount = GetLastUsedRow(permsh.Name)
        roleCount = GetLastUsedRow(roleSh.Name)
        perms = permsh.Range("A1:A" & permCount).Value
        
        For a = 1 To roleCount
            outRow = .Range("A" & .Rows.Count).End(xlUp).Row + 1
            lastRow = permCount + outRow - 1
            roleName = roleSh.Cells(a, 1).Value
            .Range("A" & outRow & ":A" & lastRow).Value = roleName
            .Range("B" & outRow & ":B" & lastRow).Value = perms
            
            For b = 1 To permCount
                checkCell = rawSh.Cells(a, 3)
                If InStr(1, checkCell, permsh.Cells(b, 1)) Then
                
            Next b
        Next a
        
    End With
    
   
End With

End Sub

Function GetLastUsedColumn(Sheet As String)
    lastCol = .Sheets(Sheet).Cells.Find(What:="*", _
            After:=Range("A1"), _
            LookAt:=xlPart, _
            LookIn:=xlFormulas, _
            SearchOrder:=xlByColumns, _
            SearchDirection:=xlPrevious, _
            MatchCase:=False).Column
    GetLastUsedColumn = lastCol
End Function

Function GetLastUsedRow(Sheet)
    lastRow = ActiveWorkbook.Sheets(Sheet).Cells.Find(What:="*", _
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

Function regex(strInput As String, matchPattern As String, Optional ByVal outputPattern As String = "$0") As Variant
    Dim inputRegexObj As New VBScript_RegExp_55.RegExp, outputRegexObj As New VBScript_RegExp_55.RegExp, outReplaceRegexObj As New VBScript_RegExp_55.RegExp
    Dim inputMatches As Object, replaceMatches As Object, replaceMatch As Object
    Dim replaceNumber As Integer

    With inputRegexObj
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = matchPattern
    End With
    With outputRegexObj
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = "\$(\d+)"
    End With
    With outReplaceRegexObj
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
    End With

    Set inputMatches = inputRegexObj.Execute(strInput)
    If inputMatches.Count = 0 Then
        regex = False
    Else
        Set replaceMatches = outputRegexObj.Execute(outputPattern)
        For Each replaceMatch In replaceMatches
            replaceNumber = replaceMatch.SubMatches(0)
            outReplaceRegexObj.Pattern = "\$" & replaceNumber

            If replaceNumber = 0 Then
                outputPattern = outReplaceRegexObj.Replace(outputPattern, inputMatches(0).Value)
            Else
                If replaceNumber > inputMatches(0).SubMatches.Count Then
                    'regex = "A to high $ tag found. Largest allowed is $" & inputMatches(0).SubMatches.Count & "."
                    regex = CVErr(xlErrValue)
                    Exit Function
                Else
                    outputPattern = outReplaceRegexObj.Replace(outputPattern, inputMatches(0).SubMatches(replaceNumber - 1))
                End If
            End If
        Next
        regex = outputPattern
    End If
End Function
