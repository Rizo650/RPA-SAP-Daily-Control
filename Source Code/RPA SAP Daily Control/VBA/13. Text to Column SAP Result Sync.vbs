Sub TextToColumnAndAddColumns()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim colC As Range

    ' Disable alerts and screen updating to prevent notifications
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    ' Set the target worksheet
    Set ws = ThisWorkbook.Sheets("SAP_Result_Sync")

    ' Find the last row with data in column C
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    ' Insert one new column after column C
    ws.Columns("D:D").Insert Shift:=xlToRight

    ' Perform Text to Columns operation on column C based on "-"
    Set colC = ws.Range("C2:C" & lastRow)
    colC.TextToColumns Destination:=ws.Range("C2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
        Tab:=False, Semicolon:=False, Comma:=False, Space:=False, _
        Other:=True, OtherChar:="-"

    ' Insert two new columns after column C (after Text to Columns)
    ws.Columns("D:E").Insert Shift:=xlToRight

    ' Re-enable screen updating and alerts after operation
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
