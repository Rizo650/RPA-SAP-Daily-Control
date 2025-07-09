Sub AddColumnsAndTextToColumnsFixedWidth()
    Dim ws As Worksheet
    Dim headerColumn As Range
    Dim lastRow As Long
    Dim headerColIndex As Integer

    ' Set worksheet sumber (ubah jika perlu)
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' Cari kolom "Document Header Text"
    Set headerColumn = ws.Rows(1).Find(What:="Document Header Text", LookIn:=xlValues, LookAt:=xlWhole)

    If Not headerColumn Is Nothing Then
        ' Dapatkan indeks kolom "Document Header Text"
        headerColIndex = headerColumn.Column

        ' Tambahkan tiga kolom setelah kolom "Document Header Text"
        ws.Columns(headerColIndex + 1).Resize(, 3).Insert Shift:=xlToRight

        ' Cari baris terakhir berdasarkan kolom "Document Header Text"
        lastRow = ws.Cells(ws.Rows.Count, headerColIndex).End(xlUp).Row

        ' Lakukan Text to Columns pada kolom "Document Header Text" menggunakan Fixed Width
        With ws.Range(ws.Cells(2, headerColIndex), ws.Cells(lastRow, headerColIndex))
            .TextToColumns _
                Destination:=ws.Cells(2, headerColIndex), _
                DataType:=xlFixedWidth, _
                FieldInfo:=Array(Array(0, 1), Array(4, 1)) ' Fixed width setelah 4 karakter (setelah "BIK-")
        End With
    End If
End Sub
