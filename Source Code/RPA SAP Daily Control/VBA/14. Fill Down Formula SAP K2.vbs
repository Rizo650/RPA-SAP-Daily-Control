Sub AddColumnsAndFormulas()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Set the target worksheet
    Set ws = ThisWorkbook.Sheets("SAP_Result_Sync")

    ' Find the last row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop through each row in column A
    For i = 2 To lastRow ' Start from row 2 to skip header
        ' Check if column A contains "Rent Room Claim"
        If ws.Cells(i, 1).Value = "Rent Room Claim" Then
            ' Set the formula for column D for the matching row
            ws.Cells(i, 4).Formula = "=B" & i & "&""-""&C" & i ' Column D is the 4th column

            ' Set the VLOOKUP formula for column E for the matching row
            ws.Cells(i, 5).Formula = "=VLOOKUP(D" & i & ", 'D:\RPA Attendance, Medical, SAP Daily Control\Source Code\RPA SAP Daily Control\[Template SAP.xlsx]BIK'!$X$2:$X$1000000, 1, FALSE)" ' Column E is the 5th column
        End If
    Next i
End Sub
