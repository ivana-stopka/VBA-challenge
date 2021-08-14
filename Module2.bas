Attribute VB_Name = "Module2"
Sub StockAnalysis2()

'''''''''''''''''''''''''
'Creating Summary Table 2
'''''''''''''''''''''''''

'Declare the variables
Dim lastrow As Integer
Dim ticker As String
Dim ws As Worksheet

'Cycle through each worksheet repeating the following
For Each ws In Worksheets

'Create the Summary Table 2 headers in bold font

ws.Cells(2, 15) = "Greatest % Increase"
ws.Cells(2, 15).Font.Bold = True
ws.Cells(3, 15) = "Greatest % Decrease"
ws.Cells(3, 15).Font.Bold = True
ws.Cells(4, 15) = "Greatest Total Volume"
ws.Cells(4, 15).Font.Bold = True
ws.Cells(1, 16) = "Ticker"
ws.Cells(1, 16).Font.Bold = True
ws.Cells(1, 17) = "Value"
ws.Cells(1, 17).Font.Bold = True

'Find the last non-blank cell in the first column of the Summary Table 2
lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

'Find and print the greatest % increase with its corresponding ticker
ws.Cells(2, 17).Value = Application.WorksheetFunction.Max(ws.Columns("K"))

For i = 2 To lastrow
    If ws.Cells(i, 11).Value = ws.Cells(2, 17).Value Then
            ticker = ws.Cells(i, 9).Value 'store ticker value
            ws.Cells(2, 16).Value = ticker 'print ticker in Summary Table 2
            ws.Range("Q2").NumberFormat = "0.00%" 'display value as a percentage
    Else
    End If
Next i

'Find and print the greatest % decrease with its corresponding ticker
ws.Cells(3, 17).Value = Application.WorksheetFunction.Min(ws.Columns("K"))

For i = 2 To lastrow
    If ws.Cells(i, 11).Value = ws.Cells(3, 17).Value Then
            ticker = ws.Cells(i, 9).Value 'store ticker value
            ws.Cells(3, 16).Value = ticker 'print ticker value in Summary Table 2
            ws.Range("Q3").NumberFormat = "0.00%" 'display value as a percentage
    Else
    End If
Next i

'Find and print the greatest total volume with its corresponding ticker
ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Columns("L"))

For i = 2 To lastrow
    If ws.Cells(i, 12).Value = ws.Cells(4, 17).Value Then
            ticker = ws.Cells(i, 9).Value 'store ticker value
            ws.Cells(4, 16).Value = ticker 'print ticker value in Summary Table 2
    Else
    End If
Next i

ws.Columns("O:Q").AutoFit 'autofit the columns to the data in Summary Table 2

Next ws

End Sub
