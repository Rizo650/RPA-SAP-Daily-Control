Sub CopyRowsToResultCheck()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRowSource As Long
    Dim lastRowTarget As Long
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
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' Find the last row in the target sheet (Result Check)
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
    
    ' Start copying to the target sheet from the row after the last used row
    targetRow = lastRowTarget + 1

    ' Loop through rows in the source sheet
    For i = 2 To lastRowSource
        cellValueA = wsSource.Cells(i, "A").Value
        cellValueC = wsSource.Cells(i, "C").Value
        cellValueE = wsSource.Cells(i, "E").Text ' Read formula result as text

        ' Skip rows where column E does not contain #N/A
        If cellValueE <> "#N/A" Then GoTo SkipRow

        ' Check conditions for Rent Room Claim
        If cellValueA = "Rent Room Claim" Then
            wsSource.Rows(i).Copy Destination:=wsTarget.Rows(targetRow)
            targetRow = targetRow + 1
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
