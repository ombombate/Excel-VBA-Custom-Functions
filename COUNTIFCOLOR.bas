Function CountIfColor(rng As Range, targetColor As Range) As Long
    Dim cell As Range
    Dim count As Long

    ' Loop through each cell in the range
    For Each cell In rng
        ' Check if the cell's color matches the target color
        If cell.Interior.Color = targetColor.Interior.Color Then
            count = count + 1
        End If
    Next cell

    ' Return the count
    CountIfColor = count
End Function
