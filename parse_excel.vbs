Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

strPath = "C:\Users\SNYOUTH\.gemini\antigravity\scratch\sample.xlsx"
Set objWorkbook = objExcel.Workbooks.Open(strPath)

' Use ADODB.Stream to write UTF-8 file
Set objStream = CreateObject("ADODB.Stream")
objStream.Type = 2    ' adTypeText
objStream.Charset = "utf-8"
objStream.Open

For Each objWorksheet In objWorkbook.Worksheets
    objStream.WriteText "--- Sheet: " & objWorksheet.Name & " ---" & vbCrLf
    
    ' Read first 3 rows
    For r = 1 To 3
        Dim rowData
        rowData = ""
        For c = 1 To 15
            Dim cellVal
            cellVal = objWorksheet.Cells(r, c).Text
            If cellVal <> "" Then
                rowData = rowData & cellVal & " | "
            End If
        Next
        If rowData <> "" Then
            objStream.WriteText "Row " & r & ": " & rowData & vbCrLf
        End If
    Next
    objStream.WriteText vbCrLf
Next

objStream.SaveToFile "C:\Users\SNYOUTH\.gemini\antigravity\scratch\excel_result.txt", 2
objStream.Close

objWorkbook.Close False
objExcel.Quit

Set objStream = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing
