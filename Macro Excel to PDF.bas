Attribute VB_Name = "Module2"
Sub ConvertExcelFilesToPDF()
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim FldrPicker As FileDialog

    
    Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)
    With FldrPicker
        .Title = "Select Folder Containing Excel Files"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub ' If the user cancels
        folderPath = .SelectedItems(1) & "\"
    End With

    ' Get the first Excel file in the folder
    fileName = Dir(folderPath & "*.xlsx")

    ' Loop through all Excel files in the folder
    Do While fileName <> ""
        ' Open the workbook
        Set wb = Workbooks.Open(folderPath & fileName)

        ' Autofit all columns and rows
        Dim ws As Worksheet
        For Each ws In wb.Worksheets
            ws.Columns.AutoFit
            ws.Rows.AutoFit
        Next ws

        ' Save as PDF
        wb.ExportAsFixedFormat Type:=xlTypePDF, fileName:=folderPath & Left(fileName, InStrRev(fileName, ".") - 1) & ".pdf"

        ' Close the workbook without saving changes
        wb.Close SaveChanges:=False

        ' Get the next Excel file in the folder
        fileName = Dir
    Loop

    MsgBox "All Excel files have been converted to PDF."
End Sub

