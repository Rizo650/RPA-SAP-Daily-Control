Sub CopyRowsToResultCheck()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim targetRow As Long
    Dim i As Long
    Dim cellValueL As String
    Dim cellValueF As String
    
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
    
    ' Start copying to the target sheet from the last row + 1
    targetRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row + 1

    ' Loop through rows in the source sheet
    For i = 2 To lastRow
        cellValueL = Trim(wsSource.Cells(i, "L").Value)
        cellValueF = Trim(wsSource.Cells(i, "F").Value)

        ' Debugging: Print the current values of F and L to Immediate Window for validation
        Debug.Print "Row " & i & ": F = '" & cellValueF & "', L = '" & cellValueL & "'"
        
        ' Check if column F contains "Document posted successfully" and column L contains "False"
        If InStr(1, cellValueF, "success", vbTextCompare) > 0 Then
            If cellValueL = "False" Then
                ' Copy the entire row to the target sheet at the next available row
                wsSource.Rows(i).Copy Destination:=wsTarget.Rows(targetRow)
                targetRow = targetRow + 1
            End If
        End If
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
