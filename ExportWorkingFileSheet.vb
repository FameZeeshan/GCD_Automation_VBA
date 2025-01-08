Sub ExportWorkingFileSheet()
    Dim ws As Worksheet
    Dim newWorkbook As Workbook
    Dim filePath As String
    
    ' Set the worksheet to be exported
    Set ws = ThisWorkbook.Sheets("Working File")
    
    ' Create a new workbook
    Set newWorkbook = Workbooks.Add
    
    ' Copy the worksheet to the new workbook
    ws.Copy Before:=newWorkbook.Sheets(1)
    
    ' Remove default sheets from the new workbook
    Application.DisplayAlerts = False
    For Each ws In newWorkbook.Sheets
        If ws.Name <> "Working File" Then ws.Delete
    Next ws
    Application.DisplayAlerts = True
    
    ' Prompt user to select the destination to save the file
    filePath = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="Save As", InitialFileName:="GCD.xlsx")
    
    ' Check if the user canceled the save dialog
    If filePath <> "False" Then
        ' Save the new workbook
        newWorkbook.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbook
        MsgBox "File saved successfully at " & filePath
    Else
        MsgBox "Save operation canceled."
    End If
    
    ' Close the new workbook
    newWorkbook.Close SaveChanges:=False
End Sub
