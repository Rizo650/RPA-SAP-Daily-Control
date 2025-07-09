Sub CopySheetToNewFile()
    Dim wbSource As Workbook
    Dim wbTarget As Workbook
    Dim targetPath As String

    ' Disable Excel notifications
    Application.DisplayAlerts = False

    ' Set source workbook
    Set wbSource = ThisWorkbook

    ' Specify target file path
    targetPath = "D:\RPA Attendance, Medical, SAP Daily Control\Source Code\RPA SAP Daily Control\Template SAP K2.xlsx"

    ' Try to open the target workbook, if not create a new one
    On Error Resume Next
    Set wbTarget = Workbooks.Open(targetPath)
    On Error GoTo 0

    ' If the target workbook doesn't exist, create a new workbook
    If wbTarget Is Nothing Then
        Set wbTarget = Workbooks.Add
    End If

    ' Copy the first sheet from source to target workbook
    wbSource.Sheets(1).Copy After:=wbTarget.Sheets(wbTarget.Sheets.Count)

    ' Save and close the target workbook
    wbTarget.SaveAs Filename:=targetPath, FileFormat:=xlOpenXMLWorkbook
    wbTarget.Close SaveChanges:=False

    ' Re-enable Excel notifications
    Application.DisplayAlerts = True
End Sub
