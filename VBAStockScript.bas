Attribute VB_Name = "Module1"
Sub VBAStocks():

'Label Column Headers
For Each ws In Worksheets

    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"

 'Set Variables
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim Volume As Double
    
    Dim StockOpen As Double
    Dim StockClose As Double
  
  'End of Loop
    Dim lastrow As Double
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


Volume = 0

'Summary Table parameters
Dim SummaryTable As Double
SummaryTable = 2

'Conditionals for Summary Table
For i = 2 To lastrow


If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
    Ticker = ws.Cells(i, 1).Value
    Volume = Volume + ws.Cells(i, 7).Value

ws.Range("I" & SummaryTable).Value = Ticker
ws.Range("L" & SummaryTable).Value = Volume

Volume = 0

StockClose = ws.Cells(i, 6)

    If StockOpen = 0 Then
    YearlyChange = 0
    PercentChange = 0
    Else:
    YearlyChange = StockClose - StockOpen
    PercentChange = YearlyChange / StockOpen
    End If
    
ws.Range("J" & SummaryTable).Value = YearlyChange
ws.Range("K" & SummaryTable).Value = PercentChange
ws.Range("K" & SummaryTable).Style = "Percent"
ws.Range("K" & SummaryTable).NumberFormat = "0.00%"

SummaryTable = SummaryTable + 1

ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
    StockOpen = ws.Cells(i, 3)
    

Else: Volume = Volume + ws.Cells(i, 7).Value

End If


    Next i


For i = 2 To lastrow

' Positive=4 , Negative=3
If ws.Range("J" & i).Value > 0 Then
        ws.Range("J" & i).Interior.ColorIndex = 4

ElseIf ws.Range("J" & i).Value < 0 Then
        ws.Range("J" & i).Interior.ColorIndex = 3
        
End If

    Next i
    
Next ws



MsgBox ("not broken yet")

End Sub
