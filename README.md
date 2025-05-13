Sub CopySheetsToNewWorkbook()
    Dim mainWB As Workbook
    Dim newWB As Workbook
    Dim ws As Worksheet
    Dim destWS As Worksheet
    Dim filePath As Variant
    Dim lastRow As Long, lastCol As Long
    Dim destLastRow As Long
    Dim isFirstSheet As Boolean
    Dim usedRange As Range

    ' Ask user to confirm the ActiveWorkbook
    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found.", vbCritical
        Exit Sub
    End If

    If MsgBox("Use '" & ActiveWorkbook.Name & "' as the main workbook?", vbYesNo + vbQuestion) = vbNo Then
        MsgBox "Please activate the correct workbook and run the macro again.", vbExclamation
        Exit Sub
    End If

    ' Set main workbook
    Set mainWB = ActiveWorkbook

    ' Ask user where to save the new .xlsb file
    filePath = Application.GetSaveAsFilename( _
                InitialFileName:="ConsolidatedData", _
                FileFilter:="Excel Binary Workbook (*.xlsb), *.xlsb", _
                Title:="Save Consolidated Workbook As")

    ' Exit if user cancels
    If filePath = False Then
        MsgBox "Operation cancelled by user.", vbExclamation
        Exit Sub
    End If

    ' Create new workbook and save as .xlsb
    Set newWB = Workbooks.Add(xlWBATWorksheet)
    Application.DisplayAlerts = False
    newWB.SaveAs Filename:=filePath, FileFormat:=xlExcel12
    Application.DisplayAlerts = True

    Set destWS = newWB.Sheets(1)
    destWS.Name = "ConsolidatedData"

    ' Loop through all worksheets in main workbook
    isFirstSheet = True
    For Each ws In mainWB.Worksheets
        With ws
            ' Get used range to calculate last row and last column correctly
            If Application.WorksheetFunction.CountA(.Cells) > 0 Then
                Set usedRange = .UsedRange
                lastRow = usedRange.Row + usedRange.Rows.Count - 1
                lastCol = usedRange.Column + usedRange.Columns.Count - 1

                ' Find last used row in destination sheet
                destLastRow = destWS.Cells(destWS.Rows.Count, 1).End(xlUp).Row
                If destLastRow = 1 And destWS.Cells(1, 1).Value = "" Then destLastRow = 0

                If isFirstSheet Then
                    ' Copy with header
                    .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).Copy _
                        Destination:=destWS.Cells(destLastRow + 1, 1)
                    isFirstSheet = False
                Else
                    ' Copy without header (starting from row 2)
                    If lastRow >= 2 Then
                        .Range(.Cells(2, 1), .Cells(lastRow, lastCol)).Copy _
                            Destination:=destWS.Cells(destLastRow + 1, 1)
                    End If
                End If
            End If
        End With
    Next ws

    ' Save and inform user
    newWB.Save
    MsgBox "Data copied successfully to: " & vbCrLf & filePath, vbInformation
End Sub


----- NEw   ---

    ' Step 5: Copy data from all sheets
    isFirstSheet = True
    destSheetIndex = 1
    Set destWS = newWB.Sheets(1)
    destWS.Name = "ConsolidatedData_" & destSheetIndex
    destLastRow = 0

    For Each ws In mainWB.Worksheets
        With ws
            If Application.WorksheetFunction.CountA(.Cells) > 0 Then
                Set usedRange = .UsedRange
                lastRow = usedRange.Row + usedRange.Rows.Count - 1
                lastCol = usedRange.Column + usedRange.Columns.Count - 1

                ' Determine the number of rows to copy
                Dim rowsToCopy As Long
                If isFirstSheet Then
                    rowsToCopy = lastRow
                Else
                    rowsToCopy = lastRow - 1 ' skip header row
                End If

                ' If the next block will exceed row limit, create a new sheet
                If destLastRow + rowsToCopy > 1048576 Then
                    destSheetIndex = destSheetIndex + 1
                    Set destWS = newWB.Sheets.Add(After:=newWB.Sheets(newWB.Sheets.Count))
                    destWS.Name = "ConsolidatedData_" & destSheetIndex
                    destLastRow = 0
                    isFirstSheet = True ' treat this as the first sheet on new destination
                End If

                ' Determine paste destination
                If isFirstSheet Then
                    .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).Copy _
                        Destination:=destWS.Cells(destLastRow + 1, 1)
                    destLastRow = destLastRow + lastRow
                    isFirstSheet = False
                Else
                    If lastRow >= 2 Then
                        .Range(.Cells(2, 1), .Cells(lastRow, lastCol)).Copy _
                            Destination:=destWS.Cells(destLastRow + 1, 1)
                        destLastRow = destLastRow + (lastRow - 1)
                    End If
                End If
            End If
        End With
    Next ws

    ' AutoFit columns on all sheets
    For Each destWS In newWB.Worksheets
        destWS.Columns.AutoFit
    Next destWS

