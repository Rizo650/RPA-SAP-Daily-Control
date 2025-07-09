Sub AddAllBordersToResultCheck()
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim targetRange As Range
    
    ' Set the target worksheet
    Set wsTarget = ThisWorkbook.Sheets("Result Check")
    
    ' Find the last row and column in the target sheet
    lastRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
    lastCol = wsTarget.Cells(1, wsTarget.Columns.Count).End(xlToLeft).Column
    
    ' Define the range to apply borders (from A1 to the last used row and column)
    Set targetRange = wsTarget.Range(wsTarget.Cells(1, 1), wsTarget.Cells(lastRow, lastCol))
    
    ' Apply all borders to the range
    With targetRange.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0 ' Black border
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
End Sub
