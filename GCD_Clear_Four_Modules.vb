Sub GCD_Clear_Four_Modules()
    Call GCD_Clear_Modules("Latest Forecast - BD WAH", "Latest Forecast - Existing WAH", "D:E,G:H", "L:N", "S:AD")
    Call GCD_Clear_Modules("Latest Forecast - BD B&M", "", "E,G:H", "L:N", "S:AD")
    Call GCD_Clear_Modules("Latest Forecast - Existing B&M", "", "E,G:H", "", "S:AD")
    MsgBox "All 4 module Data is cleared successfully!"
End Sub

Sub GCD_Clear_Modules(criteria1 As String, criteria2 As String, clearCols1 As String, clearCols2 As String, clearCols3 As String, Optional checkCriteria2 As Boolean = False)
    Dim wsWorking As Worksheet
    Dim lastRow As Long
    Dim filterRange As Range
    Dim visibleRows As Range
    Dim area As Range
    Dim clearRange1 As Range
    Dim clearRange2 As Range
    Dim clearRange3 As Range
    Dim colParts() As String
    Dim i As Integer

    ' Set reference to the "Working File" sheet
    Set wsWorking = ThisWorkbook.Sheets("Working File")

    ' Turn off screen updating to improve performance
    Application.ScreenUpdating = False
    
    ' Check if filters are already applied, remove them
    If wsWorking.AutoFilterMode Then
        wsWorking.AutoFilterMode = False
    End If

    ' Find the last row in the "Working File" sheet
    lastRow = wsWorking.Cells(wsWorking.Rows.Count, "A").End(xlUp).Row

    ' Set filter range in the "Working File" sheet (starting from row 4)
    Set filterRange = wsWorking.Range("A4:BA" & lastRow)

    ' Apply filter for the specified criteria (BA column is 53)
    filterRange.AutoFilter Field:=53, Criteria1:=criteria1
    If criteria2 <> "" Then
        filterRange.AutoFilter Field:=53, Criteria1:=criteria1, Operator:=xlOr, Criteria2:=criteria2
    End If

    ' Capture the filtered visible rows excluding the header row
    On Error Resume Next
    Set visibleRows = filterRange.Offset(1, 0).Resize(filterRange.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    ' Check if any rows are visible after applying the filter
    If Not visibleRows Is Nothing Then
        ' Combine ranges to clear content and formatting in specified columns for all visible areas
        For Each area In visibleRows.Areas
            ' Process clearCols1
            colParts = Split(clearCols1, ",")
            For i = LBound(colParts) To UBound(colParts)
                If clearRange1 Is Nothing Then
                    Set clearRange1 = area.Columns(colParts(i))
                Else
                    Set clearRange1 = Union(clearRange1, area.Columns(colParts(i)))
                End If
            Next i
            
            ' Process clearCols2
            colParts = Split(clearCols2, ",")
            For i = LBound(colParts) To UBound(colParts)
                If clearRange2 Is Nothing Then
                    Set clearRange2 = area.Columns(colParts(i))
                Else
                    Set clearRange2 = Union(clearRange2, area.Columns(colParts(i)))
                End If
            Next i

            ' Process clearCols3
            colParts = Split(clearCols3, ",")
            For i = LBound(colParts) To UBound(colParts)
                If clearRange3 Is Nothing Then
                    Set clearRange3 = area.Columns(colParts(i))
                Else
                    Set clearRange3 = Union(clearRange3, area.Columns(colParts(i)))
                End If
            Next i

            ' Additional clearing for specific criteria
            If checkCriteria2 And wsWorking.Cells(area.Row, "BA").Value = criteria2 Then
                colParts = Split(clearCols1 & "," & clearCols3, ",")
                For i = LBound(colParts) To UBound(colParts)
                    If clearRange1 Is Nothing Then
                        Set clearRange1 = area.Columns(colParts(i))
                    Else
                        Set clearRange1 = Union(clearRange1, area.Columns(colParts(i)))
                    End If
                Next i
            End If
        Next area
        
        ' Clear contents and background color in the combined ranges
        If Not clearRange1 Is Nothing Then
            clearRange1.ClearContents
            clearRange1.Interior.ColorIndex = xlNone
        End If
        
        If Not clearRange2 Is Nothing Then
            clearRange2.ClearContents
            clearRange2.Interior.ColorIndex = xlNone
        End If
        
        If Not clearRange3 Is Nothing Then
            clearRange3.ClearContents
            clearRange3.Interior.ColorIndex = xlNone
        End If
    End If

    ' Remove filter
    wsWorking.AutoFilterMode = False

    ' Turn screen updating back on
    Application.ScreenUpdating = True
End Sub