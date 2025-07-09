Sub ConvertToValidString()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    
    ' Daftar nama sheet dan kolom yang akan diubah
    Dim sheetData As Variant
    sheetData = Array( _
        Array("PDDN", "Y"), _
        Array("COMMON", "Y"), _
        Array("MEDICAL", "Y"), _
        Array("TRANSPORT", "Y"), _
        Array("BIK", "X") _
    )
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error Resume Next ' Handle error jika sheet tidak ada
    For Each ws In ThisWorkbook.Worksheets
        For i = LBound(sheetData) To UBound(sheetData)
            If ws.Name = sheetData(i)(0) Then
                ' Dapatkan kolom yang ditentukan
                lastRow = ws.Cells(ws.Rows.Count, sheetData(i)(1)).End(xlUp).Row
                For Each cell In ws.Range(sheetData(i)(1) & "2:" & sheetData(i)(1) & lastRow)
                    ' Ubah nilai menjadi string, terlepas dari tipe data
                    cell.Value = "'" & CStr(cell.Value)
                Next cell
            End If
        Next i
    Next ws
    On Error GoTo 0 ' Matikan error handler
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
