Sub ExportAllCellsToCSV()
    Dim ws As Worksheet
    Dim csvFile As String
    Dim cell As Range
    Dim rowNum As Long
    Dim colNum As Long
    Dim csvLine As String
    Dim fso As Object
    Dim ts As Object

    ' Set the CSV file path
    csvFile = "C:\Path\To\Your\ExportedFile.csv"

    ' Create a FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Create a TextStream object to write to the CSV file
    Set ts = fso.CreateTextFile(csvFile, True)

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Loop through each row in the worksheet
        For rowNum = 1 To ws.UsedRange.Rows.Count
            csvLine = ""
            ' Loop through each column in the row
            For colNum = 1 To ws.UsedRange.Columns.Count
                ' Get the cell value and add it to the CSV line
                csvLine = csvLine & ws.Cells(rowNum, colNum).Value
                ' Add a comma if it's not the last column
                If colNum < ws.UsedRange.Columns.Count Then
                    csvLine = csvLine & ","
                End If
            Next colNum
            ' Write the CSV line to the file
            ts.WriteLine csvLine
        Next rowNum
    Next ws

    ' Close the TextStream object
    ts.Close

    ' Inform the user that the export is complete
    MsgBox "Export complete! The CSV file is saved at: " & csvFile
End Sub



USE YourDatabaseName;
GO
CREATE DATABASE AUDIT SPECIFICATION MyDatabaseAuditSpec
FOR SERVER AUDIT MyAudit
ADD (SELECT ON SCHEMA::dbo BY public);
GO
