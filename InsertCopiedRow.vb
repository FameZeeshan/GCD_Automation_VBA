Sub InsertCopiedRow()
    Dim targetRow As Long
    Dim numRows As Long
    Dim copiedRange As Range
    Dim ws As Worksheet
   
    ' Reference the "Working File" sheet
    Set ws = ThisWorkbook.Sheets("Working File")
   
    ' Prompt user for the row number to add additional rows
    targetRow = InputBox("Enter the row number where you want to add additional rows:", "Insert Rows", 1)
   
    ' Prompt user for number of rows to add
    numRows = InputBox("Enter the number of additional rows to add:", "Insert Rows", 1)
   
    ' Validate user input
    If targetRow < 1 Or numRows < 1 Then
        MsgBox "Invalid input. Please enter valid row number and number of rows.", vbExclamation
        Exit Sub
    End If
   
    ' Set the copied range (assuming a single row is copied)
    Set copiedRange = ws.Rows(targetRow)
   
    ' Insert additional rows below the target row
    ws.Rows(targetRow + 1).Resize(numRows).Insert Shift:=xlDown
   
    ' Paste the copied cells
    copiedRange.Copy
    ws.Rows(targetRow + 1).Resize(numRows).PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False ' Clear the clipboard
   
    ' Drag down formula from the specified row
    With ws
        .Range(.Cells(targetRow, 1), .Cells(targetRow + numRows, 1)).FillDown
    End With
End Sub


