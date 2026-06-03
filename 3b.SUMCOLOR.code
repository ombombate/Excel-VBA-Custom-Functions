Function SumColor(rng As Range) As Double
    Dim cell As Range
    Dim cellValue As Double

    For Each cell In rng

        Select Case cell.Interior.ColorIndex

            Case 6
                cellValue = 5

            Case 3
                cellValue = 0

            Case 43
                cellValue = 10

            Case 56
                cellValue = -5

            Case 33
                cellValue = 15

            Case Else
                cellValue = 0

        End Select

        SumColor = SumColor + cellValue

    Next cell

End Function
