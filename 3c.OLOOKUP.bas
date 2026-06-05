Function OLOOKUP(LookupValue As Variant, _
                 TableRange As Range, _
                 ReturnColumn As String, _
                 Optional ExactMatch As Boolean = True)

    Dim LookupCol As Range
    Dim ReturnCol As Range
    Dim ColNum As Long
    Dim ResultRow As Variant

    ReturnColumn = UCase(ReturnColumn)

    ColNum = Range(ReturnColumn & "1").Column

    Set LookupCol = TableRange.Columns(1)

    Set ReturnCol = Columns(ColNum)

    If ExactMatch Then

        ResultRow = Application.Match(LookupValue, LookupCol, 0)

    Else

        ResultRow = Application.Match(LookupValue, LookupCol, 1)

    End If

    If IsError(ResultRow) Then

        OLOOKUP = "Not Found"

    Else

        OLOOKUP = ReturnCol.Cells(ResultRow).Value

    End If

End Function
