Sub DeleteSpecifiedRows()
    Dim targetRow As Long
    Dim numRows As Long
    Dim ws As Worksheet
    
    ' Reference the "Working File" sheet
    Set ws = ThisWorkbook.Sheets("Working File")
    
    ' Prompt user for the row number to start deleting rows
    targetRow = InputBox("Enter the row number below which you want to delete rows:", "Delete Rows", 1)
    
    ' Prompt user for number of rows to delete
    numRows = InputBox("Enter the number of rows to delete:", "Delete Rows", 1)
    
    ' Validate user input
    If targetRow < 1 Or numRows < 1 Then
        MsgBox "Invalid input. Please enter valid row number and number of rows.", vbExclamation
        Exit Sub
    End If
    
    ' Delete the specified number of rows below the target row
    ws.Rows(targetRow + 1 & ":" & targetRow + numRows).Delete Shift:=xlUp
End Sub
