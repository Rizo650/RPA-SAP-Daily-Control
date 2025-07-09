Sub CopyBIKRows()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim headerColIndex As Integer
    Dim headerValue As String
    Dim targetRow As Long
    Dim copyRange As Range

    ' Set worksheet sumber dan target
    Set wsSource = ThisWorkbook.Sheets("Sheet1") ' Ganti dengan nama sheet sumber
    Set wsTarget = ThisWorkbook.Sheets("BIK")  ' Ganti dengan nama sheet target

    ' Cari kolom "Document Header Text" di Sheet1
    On Error Resume Next
    headerColIndex = wsSource.Rows(1).Find(What:="Document Header Text", LookIn:=xlValues, LookAt:=xlWhole).Column
    On Error GoTo 0

    ' Jika kolom tidak ditemukan
    If headerColIndex = 0 Then
        MsgBox "Kolom 'Document Header Text' tidak ditemukan!"
        Exit Sub
    End If

    ' Cari baris terakhir berdasarkan kolom "Document Header Text"
    lastRow = wsSource.Cells(wsSource.Rows.Count, headerColIndex).End(xlUp).Row

    ' Tentukan baris pertama yang kosong di sheet BIK
    targetRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row + 1

    ' Menonaktifkan screen updating dan perhitungan untuk mempercepat eksekusi
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Loop untuk memeriksa setiap baris di kolom "Document Header Text"
    For i = 2 To lastRow
        headerValue = wsSource.Cells(i, headerColIndex).Value

        ' Cek apakah nilai di kolom "Document Header Text" mengandung "BIK-"
        If InStr(headerValue, "BIK-") > 0 Then
            ' Menambahkan baris yang sesuai ke dalam range untuk disalin
            If copyRange Is Nothing Then
                Set copyRange = wsSource.Rows(i)
            Else
                Set copyRange = Union(copyRange, wsSource.Rows(i))
            End If
        End If
    Next i

    ' Jika ada baris yang cocok, salin semua sekaligus
    If Not copyRange Is Nothing Then
        copyRange.Copy Destination:=wsTarget.Cells(targetRow, 1)
    End If

    ' Mengaktifkan kembali screen updating dan perhitungan
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
