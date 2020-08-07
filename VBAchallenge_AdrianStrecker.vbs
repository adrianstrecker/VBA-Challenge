Attribute VB_Name = "Module1"
Sub main()

'set variables
Set ws = ActiveSheet

For Each ws In Worksheets
'Add colummn headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
Dim summaryRow As Integer
    summaryRow = 2
    
Dim tickertotal As Double
Dim symbol As String
Dim yrchange, percentChange, openPrice, closePrice As Double
'Initialize tickertotal to 0
 tickertotal = 0
 openPrice = ws.Cells(2, 3).Value
'create summary table

For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
'ws.Cells.SpecialCells(xlCellTypeLastCell).Row
'Loop code
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    symbol = ws.Cells(i, 1).Value
    tickertotal = tickertotal + ws.Cells(i, 7).Value
    yrchange = ws.Cells(i, 6) - openPrice
        If openPrice = 0 Then
         percentChange = 0
        Else
         percentChange = yrchange / openPrice
        End If
        'Color Conditional formatting
      
      ws.Range("I" & summaryRow).Value = symbol
      ws.Range("L" & summaryRow).Value = tickertotal
      ws.Range("J" & summaryRow).Value = yrchange
      ws.Range("K" & summaryRow).Value = percentChange
        'reset tickertotal
         tickertotal = 0
    'incrementing summaryRow & openPrice
     summaryRow = summaryRow + 1
     openPrice = ws.Cells(i + 1, 3).Value
    
    Else
         yrchange = ws.Cells(i, 6) - openPrice
        tickertotal = tickertotal + Cells(i, 7).Value
    
          If ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = "3"
          Else
        ws.Cells(i, 10).Interior.ColorIndex = "4"
         End If
         

'Format Cells assistance from docs.microsoft.com/en-us/office/vba/api/excel.range.numberformat
    ws.Cells(i, 11).NumberFormat = "0.00%"

    End If
            

Next i

For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
'Color Conditional formatting
      If ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = "3"
          Else
        ws.Cells(i, 10).Interior.ColorIndex = "4"
         End If

Next i

Next ws

'Add headers for extra calculations
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"

For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
'Color Conditional formatting
      If Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.ColorIndex = "3"
          Else
        Cells(i, 10).Interior.ColorIndex = "4"
         End If
     
'Format to percentage
    Cells(i, 11).NumberFormat = "0.00%"

Next i
Cells(2, 16).NumberFormat = "0.00%"
Cells(3, 16).NumberFormat = "0.00%"
'Help from VBA for Dummies page 125 and mrexcel.com/excel-tips/excel-lookup-index-matc
'Greatest % increase
Cells(2, 16).Formula = "=MAX(K:K)"
Cells(2, 15).Formula = "=INDEX(I:I, MATCH(P2, K:K, FALSE), 1)"

'Greatest % decrease
Cells(3, 16).Formula = "=MIN(K:K)"
Cells(3, 15).Formula = "=INDEX(I:I, MATCH(P3, K:K, FALSE), 1)"

'Greatest Total Stock Vol
Cells(4, 16).Formula = "=MAX(L:L)"
Cells(4, 15).Formula = "=INDEX(I:I, MATCH(P4, L:L, FALSE), 1)"

End Sub
