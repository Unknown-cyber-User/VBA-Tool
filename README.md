Sub CopySheetsToNewWorkbook()
    Dim mainWB As Workbook
    Dim newWB As Workbook
    Dim ws As Worksheet
    Dim destWS As Worksheet
    Dim filePath As Variant
    Dim lastRow As Long, lastCol As Long
    Dim destLastRow As Long
    Dim isFirstSheet As Boolean

    ' Set main workbook
    Set mainWB = ThisWorkbook

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
    Application.DisplayAlerts = False ' To overwrite without prompt (if needed)
    newWB.SaveAs Filename:=filePath, FileFormat:=xlExcel12 ' xlExcel12 = .xlsb
    Application.DisplayAlerts = True

    Set destWS = newWB.Sheets(1)
    destWS.Name = "ConsolidatedData"

    ' Loop through all worksheets in main workbook
    isFirstSheet = True
    For Each ws In mainWB.Worksheets
        With ws
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column

            If lastRow > 1 Then
                ' Find last used row in destination sheet
                destLastRow = destWS.Cells(destWS.Rows.Count, 1).End(xlUp).Row
                If destLastRow = 1 And destWS.Cells(1, 1).Value = "" Then destLastRow = 0

                If isFirstSheet Then
                    ' Copy with header
                    .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).Copy _
                        Destination:=destWS.Cells(destLastRow + 1, 1)
                    isFirstSheet = False
                Else
                    ' Copy without header
                    .Range(.Cells(2, 1), .Cells(lastRow, lastCol)).Copy _
                        Destination:=destWS.Cells(destLastRow + 1, 1)
                End If
            End If
        End With
    Next ws

    ' Save and inform user
    newWB.Save
    MsgBox "Data copied successfully to: " & vbCrLf & filePath, vbInformation
End Sub
