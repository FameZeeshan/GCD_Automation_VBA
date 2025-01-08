Sub CopyPasteFourModuleData()
    ' Handle the data copy-paste operation for all forecast types
    Dim wsWorking As Worksheet
    Dim wsSource As Worksheet
    Dim lastRowSource As Long
    Dim lastRowDest As Long
    Dim filterRange As Range
    Dim visibleRows As Range
    Dim area As Range
    Dim startRow As Long
    Dim endRow As Long
    Dim forecastTypes As Variant
    Dim forecastType As Variant
    Dim srcData1 As Variant
    Dim srcData2 As Variant
    Dim srcData3 As Variant
    Dim srcData4 As Variant
    Dim srcData5 As Variant
    Dim srcData6 As Variant

    ' Turn off screen updating and automatic calculation for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Set reference to the "Working File" sheet
    Set wsWorking = ThisWorkbook.Sheets("Working File")

    ' Define the forecast types
    forecastTypes = Array("Latest Forecast - BD B&M", "Latest Forecast - BD WAH", "Latest Forecast - Existing WAH", "Latest Forecast - Existing B&M")

    ' Loop through each forecast type
    For Each forecastType In forecastTypes
        ' Set reference to the source sheet
        Set wsSource = ThisWorkbook.Sheets(forecastType)

        ' Find the last row in the source sheet
        lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

        ' Find the last row in the destination sheet
        lastRowDest = wsWorking.Cells(wsWorking.Rows.Count, "A").End(xlUp).Row

        ' Set the filter range in the "Working File" sheet (starting from row 4)
        Set filterRange = wsWorking.Range("A4:BA" & lastRowDest)

        ' Apply filter for the specified forecast type in BA column (Field 53 refers to BA column)
        filterRange.AutoFilter Field:=53, Criteria1:=forecastType

        ' Capture the filtered visible rows
        On Error Resume Next
        Set visibleRows = filterRange.Offset(1, 0).Resize(filterRange.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        ' Check if any rows are visible after applying the filter
        If Not visibleRows Is Nothing Then
            ' Load data from the source sheet into arrays
            srcData1 = wsSource.Range("A2:A" & lastRowSource).Value
            srcData2 = wsSource.Range("B2:B" & lastRowSource).Value
            srcData3 = wsSource.Range("C2:C" & lastRowSource).Value
            srcData4 = wsSource.Range("D2:D" & lastRowSource).Value
            srcData5 = wsSource.Range("E2:G" & lastRowSource).Value
            srcData6 = wsSource.Range("H2:S" & lastRowSource).Value

            ' Loop through each visible row area
            For Each area In visibleRows.Areas
                startRow = area.Row
                endRow = area.Rows.Count + startRow - 1

                ' Write data to the "Working File" sheet based on forecast type
                Select Case forecastType
                    Case "Latest Forecast - BD B&M"
                        wsWorking.Range("E" & startRow & ":E" & endRow).Value = srcData2
                        wsWorking.Range("G" & startRow & ":G" & endRow).Value = srcData3
                        wsWorking.Range("H" & startRow & ":H" & endRow).Value = srcData4
                        wsWorking.Range("L" & startRow & ":N" & endRow).Value = srcData5
                        wsWorking.Range("S" & startRow & ":AD" & endRow).Value = srcData6
                    Case "Latest Forecast - BD WAH", "Latest Forecast - Existing WAH"
                        wsWorking.Range("D" & startRow & ":D" & endRow).Value = srcData1
                        wsWorking.Range("E" & startRow & ":E" & endRow).Value = srcData2
                        wsWorking.Range("G" & startRow & ":G" & endRow).Value = srcData3
                        wsWorking.Range("H" & startRow & ":H" & endRow).Value = srcData4
                        wsWorking.Range("L" & startRow & ":N" & endRow).Value = srcData5
                        wsWorking.Range("S" & startRow & ":AD" & endRow).Value = srcData6
                    Case "Latest Forecast - Existing B&M"
                        wsWorking.Range("E" & startRow & ":E" & endRow).Value = srcData2
                        wsWorking.Range("G" & startRow & ":G" & endRow).Value = srcData3
                        wsWorking.Range("H" & startRow & ":H" & endRow).Value = srcData4
                        wsWorking.Range("S" & startRow & ":AD" & endRow).Value = srcData6
                End Select
            Next area
        End If

        ' Remove filter
        wsWorking.AutoFilterMode = False
    Next forecastType

    ' Turn on screen updating and automatic calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' Display a message indicating the completion
    MsgBox "Data for all forecast types is pasted in Working File"
End Sub