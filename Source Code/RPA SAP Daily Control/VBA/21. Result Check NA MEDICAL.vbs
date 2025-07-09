Sub CopyRowsToResultCheck()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim targetRow As Long
    Dim sourceData As Variant
    Dim targetData As Variant
    Dim i As Long, j As Long
    Dim rowCount As Long

    ' Disable alerts and events to prevent interruptions
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    On Error GoTo ErrorHandler

    ' Set the source and target worksheets
    Set wsSource = ThisWorkbook.Sheets("SAP_Result_Sync")
    Set wsTarget = ThisWorkbook.Sheets("Result Check")
    
    ' Find the last row in the source sheet
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' Load the source data into an array
    sourceData = wsSource.Range("A2:E" & lastRow).Value

    ' Initialize the target row starting position
    targetRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row + 1

    ' Initialize the target array (with maximum possible size)
    ReDim targetData(1 To 1, 1 To 5) ' Assuming 5 columns to copy

    rowCount = 0 ' To track how many rows to copy

    ' Loop through the source data (start from row 2)
    For i = 1 To UBound(sourceData, 1)
        If sourceData(i, 1) = "Medical Claim" And IsError(sourceData(i, 5)) Then
            If sourceData(i, 5) = CVErr(xlErrNA) Then
                rowCount = rowCount + 1

                ' Expand targetData array to hold the new row
                ReDim Preserve targetData(1 To rowCount, 1 To 5)

                ' Copy the row from sourceData to targetData
                For j = 1 To 5
                    targetData(rowCount, j) = sourceData(i, j)
                Next j
            End If
        End If
    Next i

    ' Write the targetData array to the target sheet starting from targetRow
    If rowCount > 0 Then
        wsTarget.Range(wsTarget.Cells(targetRow, 1), wsTarget.Cells(targetRow + rowCount - 1, 5)).Value = targetData
    End If

    ' Reactivate alerts and events
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Exit Sub

ErrorHandler:
    ' Handle errors gracefully
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
