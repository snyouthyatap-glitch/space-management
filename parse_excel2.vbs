Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

strPath = "C:\Users\SNYOUTH\.gemini\antigravity\scratch\sample.xlsx"
Set objWorkbook = objExcel.Workbooks.Open(strPath)

Set objStream = CreateObject("ADODB.Stream")
objStream.Type = 2    ' adTypeText
objStream.Charset = "utf-8"
objStream.Open

For Each objWorksheet In objWorkbook.Worksheets
    objStream.WriteText "=== Sheet: " & objWorksheet.Name & " ===" & vbCrLf
    
    ' Read first 5 rows and 20 columns
    For r = 1 To 5
        Dim rowData
        rowData = ""
        For c = 1 To 20
            Dim cellVal
            cellVal = objWorksheet.Cells(r, c).Text
            If cellVal = "" Then
                cellVal = "EMPTY"
            End If
            rowData = rowData & cellVal & " | "
        Next
        objStream.WriteText "Row " & r & ": " & rowData & vbCrLf
    Next
    objStream.WriteText vbCrLf
Next

objStream.SaveToFile "C:\Users\SNYOUTH\.gemini\antigravity\scratch\excel_headers.txt", 2
objStream.Close

objWorkbook.Close False
objExcel.Quit

Set objStream = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing
