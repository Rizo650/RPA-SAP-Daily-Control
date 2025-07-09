Sub TextToColumnMEDICAL()
    Dim ws As Worksheet
    Dim headerColumn As Range
    Dim headerColIndex As Integer
    Dim lastRow As Long

    ' Set worksheet PDDN
    Set ws = ThisWorkbook.Sheets("MEDICAL")

    ' Cari kolom "Document Header Text" di baris pertama
    Set headerColumn = ws.Rows(1).Find(What:="Document Header Text", LookIn:=xlValues, LookAt:=xlWhole)

    ' Jika kolom "Document Header Text" ditemukan
    If Not headerColumn Is Nothing Then
        ' Dapatkan indeks kolom "Document Header Text"
        headerColIndex = headerColumn.Column

        ' Tentukan baris terakhir di kolom "Document Header Text"
        lastRow = ws.Cells(ws.Rows.Count, headerColIndex).End(xlUp).Row

        ' Lakukan Text to Columns setelah kolom "Document Header Text" berdasarkan simbol "-"
        With ws.Range(ws.Cells(2, headerColIndex + 1), ws.Cells(lastRow, headerColIndex + 1))
            .TextToColumns _
                Destination:=ws.Cells(2, headerColIndex + 1), _
                DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, _
                Semicolon:=False, _
                Comma:=False, _
                Space:=False, _
                Other:=True, _
                OtherChar:="-"
        End With
    Else
        MsgBox "'Document Header Text' tidak ditemukan!"
    End If
End Sub
