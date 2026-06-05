Function OLOOKUP(LookupValue As Variant, _
                 TableRange As Range, _
                 HeaderRef As Variant) As Variant

    Dim HeaderCell As Range
    Dim LookupRange As Range
    Dim HeaderName As String
    Dim ReturnCol As Long
    Dim FoundRow As Variant

    HeaderName = CStr(HeaderRef)

    ReturnCol = 0

    For Each HeaderCell In TableRange.Rows(1).Cells

        If Trim(UCase(HeaderCell.Value)) = Trim(UCase(HeaderName)) Then

            ReturnCol = HeaderCell.Column - TableRange.Columns(1).Column + 1
            Exit For

        End If

    Next HeaderCell

    If ReturnCol = 0 Then
        OLOOKUP = "Header Not Found"
        Exit Function
    End If

    Set LookupRange = TableRange.Columns(1).Offset(1, 0).Resize(TableRange.Rows.Count - 1)

    FoundRow = Application.Match(LookupValue, LookupRange, 0)

    If IsError(FoundRow) Then
        OLOOKUP = "Value Not Found"
        Exit Function
    End If

    OLOOKUP = TableRange.Cells(FoundRow + 1, ReturnCol).Value

End Function
