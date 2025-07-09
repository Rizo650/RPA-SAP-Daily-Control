Sub AddVLookupFormulaIfPDDNClaim()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim formula As String

    ' Set the target worksheet
    Set ws = ThisWorkbook.Sheets("SAP_Result_Sync")

    ' Find the last row in column A with "Rent Room Claim" or "PDDN Claim"
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop through each row in column A
    For i = 2 To lastRow ' Start from row 2 to skip header
        ' Check if column A contains "PDDN Claim"
        If ws.Cells(i, 1).Value = "PDDN Claim" Then
            ' Set the VLOOKUP formula in column E for the matching row
            formula = "=VLOOKUP(B" & i & ", '[Template SAP.xlsx]PDDN'!$Y$2:$Y$1000000, 1, FALSE)"
            ws.Cells(i, 5).Formula = formula ' Column E is the 5th column
        End If
    Next i
End Sub
