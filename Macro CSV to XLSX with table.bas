Attribute VB_Name = "Module1"
Sub CSV_to_XLSX()
    Dim xFd As FileDialog
    Dim xSPath As String
    Dim xCSVFile As String
    Dim xWsheet As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rangeAddress As String
    Dim orangeColor As Long

    Application.DisplayAlerts = False
    Application.StatusBar = True
    xWsheet = ActiveWorkbook.Name
    Set xFd = Application.FileDialog(msoFileDialogFolderPicker)
    xFd.Title = "Select a folder:"

    If xFd.Show = -1 Then
        xSPath = xFd.SelectedItems(1)
    Else
        Exit Sub
    End If

    If Right(xSPath, 1) <> "\" Then xSPath = xSPath + "\"
    xCSVFile = Dir(xSPath & "*.csv")

    Do While xCSVFile <> ""
        Application.StatusBar = "Converting: " & xCSVFile
        Workbooks.Open Filename:=xSPath & xCSVFile
        Set ws = ActiveSheet

        ' Save the workbook as .xlsx
        ActiveWorkbook.SaveAs Replace(xSPath & xCSVFile, ".csv", ".xlsx", vbTextCompare), xlWorkbookDefault

        ' Find the last row and column
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

        ' Define the range address
        rangeAddress = ws.Cells(1, 1).Address & ":" & ws.Cells(lastRow, lastCol).Address

        ' Create the table
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(rangeAddress), , xlYes)
        tbl.Name = "MyTable"

        ' Apply the table style
        tbl.TableStyle = "TableStyleMedium3"

        ' Set header row color to orange
        orangeColor = RGB(255, 165, 0)
        tbl.HeaderRowRange.Interior.Color = orangeColor

        ' Close the workbook
        ActiveWorkbook.Close SaveChanges:=True

        ' Move to the next CSV file
        xCSVFile = Dir
    Loop

    Application.StatusBar = False
    Application.DisplayAlerts = True
End Sub
