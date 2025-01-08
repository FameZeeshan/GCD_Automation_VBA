Sub DeleteWorkSheets()
    Dim sheetNames As Variant
    Dim sheetName As Variant
    Dim ws As Worksheet
    Dim wb As Workbook
    
    ' Array of sheet names to delete
    sheetNames = Array("Latest Forecast - Existing B&M", "Latest Forecast - BD B&M", "Latest Forecast - BD WAH", "Latest Forecast - Existing WAH")
    
    ' Reference the workbook
    Set wb = ThisWorkbook
    
    ' Loop through each sheet name in the array
    For Each sheetName In sheetNames
        ' Check if the sheet exists
        On Error Resume Next
        Set ws = wb.Sheets(sheetName)
        On Error GoTo 0
        
        ' If the sheet exists, delete it
        If Not ws Is Nothing Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            Set ws = Nothing
        End If
    Next sheetName
    
    MsgBox "Specified unwanted sheets have been deleted if they existed."
End Sub
