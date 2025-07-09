Sub CopyRowsToResultCheck()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim targetRow As Long
    Dim i As Long
    Dim cellValueA As String
    Dim cellValueC As String
    Dim cellValueE As String
    
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
    
    ' Find the last row in the target sheet and start copying after it
    targetRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row + 1

    ' Loop through rows in the source sheet
    For i = 2 To lastRow
        cellValueA = wsSource.Cells(i, "A").Value
        cellValueC = wsSource.Cells(i, "C").Value
        cellValueE = wsSource.Cells(i, "E").Text ' Read formula result as text

        ' Skip rows where column E does not contain #N/A
        If cellValueE <> "#N/A" Then GoTo SkipRow

        ' Check conditions for Transportation Claim
        If cellValueA = "Transportation Claim" Then
            If Not (InStr(cellValueC, "1100") > 0 Or InStr(cellValueC, "3100") > 0) Then
                wsSource.Rows(i).Copy Destination:=wsTarget.Rows(targetRow)
                targetRow = targetRow + 1
            End If
        End If
SkipRow:
    Next i

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
