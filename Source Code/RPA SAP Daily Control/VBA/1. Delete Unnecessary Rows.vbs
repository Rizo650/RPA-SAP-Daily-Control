Sub CopyDataToNewFile()
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wbTarget As Workbook
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim sourceRange As Range
    Dim targetPath As String

    ' Set workbook dan worksheet sumber
    Set wbSource = ThisWorkbook
    Set wsSource = wbSource.Sheets(1) ' Ganti sesuai nama sheet sumber

    ' Cari baris terakhir di kolom D
    lastRow = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row

    ' Tentukan range yang akan disalin
    Set sourceRange = wsSource.Range("D9:AF" & lastRow)

    ' Tentukan path untuk file template yang sudah ada
    targetPath = "D:\RPA Attendance, Medical, SAP Daily Control\Source Code\RPA SAP Daily Control\Template SAP.xlsx"

    ' Buka workbook template yang sudah ada
    Set wbTarget = Workbooks.Open(targetPath)
    Set wsTarget = wbTarget.Sheets(1) ' Ganti dengan nama sheet target yang sesuai

    ' Salin data ke workbook baru
    sourceRange.Copy
    wsTarget.Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats

    ' Simpan workbook baru (Template SAP.xlsx sudah ada, jadi kita simpan dengan perubahan)
    Application.DisplayAlerts = False
    wbTarget.Save
    Application.DisplayAlerts = True

    ' Tutup workbook baru
    wbTarget.Close SaveChanges:=False

End Sub
