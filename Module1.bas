Attribute VB_Name = "Module1"
Sub StockAnalysis()

''''''''''''''''''''''''''
'Creating Summary Table 1
''''''''''''''''''''''''''

'Declare the variables
Dim ws As Worksheet
Dim lastrow As Long
Dim lastcolumn As Long
Dim ticker As String
Dim opening As Double
Dim closing As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim stock_counter As Double
Dim summary_table_row As Integer

'Cycle through each worksheet repeating the following
For Each ws In Worksheets

'Find the last non-blank cell in the first column
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Find the last non-blank cell in the first row
lastcolumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

'Create the summary table headers in bold font
ws.Cells(1, lastcolumn + 2) = "Ticker"
ws.Cells(1, lastcolumn + 2).Font.Bold = True
ws.Cells(1, lastcolumn + 3) = "Yearly Change"
ws.Cells(1, lastcolumn + 3).Font.Bold = True
ws.Cells(1, lastcolumn + 4) = "Percent Change"
ws.Cells(1, lastcolumn + 4).Font.Bold = True
ws.Cells(1, lastcolumn + 5) = "Total Stock Volume"
ws.Cells(1, lastcolumn + 5).Font.Bold = True

summary_table_row = 2 'specify starting row in summary table
opening = ws.Cells(2, 3).Value 'store first opening value
stock_counter = ws.Cells(2, 7).Value 'store first stock count
    
    For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value 'store ticker value
                closing = ws.Cells(i, 6).Value 'store closing balance
                yearly_change = closing - opening 'calculate yearly change
                
                    If opening = 0 Then 'calculate percent change but only if opening is not zero, or will get 0 divide by 0 resulting in overflow
                        percent_change = closing
                    Else
                        percent_change = (closing - opening) / opening
                    End If
                    
                ws.Range("I" & summary_table_row).Value = ticker 'print ticker in summary table
                ws.Range("J" & summary_table_row).Value = yearly_change 'print yearly change in summary table
                
                    If ws.Range("J" & summary_table_row).Value >= 0 Then
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 4 'if yearly change is equal to or greater than zero, color cell interior green
                    Else
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 3 'otherwise its negative, so colour cell interior red
                    End If
                
                ws.Range("K" & summary_table_row).Value = percent_change 'print percent change in summary table
                ws.Range("K" & summary_table_row).NumberFormat = "0.00%" 'display value as a percentage
                
                stock_counter = stock_counter + ws.Cells(i, 7).Value 'add final stock to stock counter
                ws.Range("L" & summary_table_row).Value = stock_counter 'print stock total in summary table
                
                summary_table_row = summary_table_row + 1 'add a row to the summary table
                closing = 0 'reset the closing value
                opening = ws.Cells(i + 1, 3).Value 'store new opening balance for next ticker
                stock_counter = ws.Cells(i + 1, 7).Value 'store new stock balance for next ticker
                                
            Else
                stock_counter = stock_counter + ws.Cells(i, 7).Value
                
            End If
            
        
    Next i
    
ws.Columns("I:L").AutoFit 'autofit the columns to the data in the summary table

Next ws


End Sub
