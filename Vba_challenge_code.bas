Attribute VB_Name = "Module1"
 Sub Ticker()

Dim rowcount As LongLong
Dim tickercount As Long
Dim ws As Worksheet
Dim total_tickers As Integer
Dim increase_val As Double
Dim decrease_val As Double
Dim max_val As LongLong
Dim ticker_max As String
Dim ticker_min As String
Dim ticker_total As String
Dim open_val As Double
Dim closing_item As Double
Dim total_volume As LongLong



For Each ws In Worksheets

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Toal Stock Volume"
    
    rowcount = ws.Cells(Rows.Count, 1).End(xlUp).Row
    tickercount = 2
    total_tickers = 0
    
    open_val = ws.Cells(2, 3).Value
    total_volume = 0
    
    For i = 2 To rowcount
   
     'closing_item = ws.Cells(i, 7).Value
     If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1) Then
        total_volume = total_volume + ws.Cells(i, 7).Value
     Else
            ws.Cells(tickercount, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(tickercount, 10).Value = ws.Cells(i, 6).Value - open_val
                  If (ws.Cells(tickercount, 10).Value > 0) Then
                        ws.Cells(tickercount, 10).Interior.ColorIndex = 4
                  Else
                        ws.Cells(tickercount, 10).Interior.ColorIndex = 3
                 End If
                 
        ws.Cells(tickercount, 11).Value = (ws.Cells(i, 6).Value / open_val) - 1
        ws.Cells(tickercount, 11).NumberFormat = "0.00%"
        ws.Cells(tickercount, 12).Value = total_volume
        ws.Cells(tickercount, 12).NumberFormat = "General"
        
        open_val = ws.Cells(i + 1, 3).Value
        total_volume = 0
        total_tickers = total_tickers + 1
        tickercount = tickercount + 1
        End If
        rowcount = rowcount + 1
        
    Next i

    ws.Cells(2, 15).Value = "Greatest % increase"
    ws.Cells(3, 15).Value = "Greatest % decrease"
    ws.Cells(4, 15).Value = "Greatest total volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"


    increase_val = 0
    decrease_val = 0
    max_val = 0
    
    
    For i = 2 To total_tickers
                   
            If ws.Cells(i, 11).Value > increase_val Then
                increase_val = ws.Cells(i, 11).Value
                ticker_max = ws.Cells(i, 9).Value
            
            ElseIf ws.Cells(i, 11).Value < decrease_val Then
                decrease_val = ws.Cells(i, 11).Value
                ticker_min = ws.Cells(i, 9).Value
            End If
           If ws.Cells(i, 12).Value > max_val Then
                max_val = ws.Cells(i, 12).Value
                ticker_total = ws.Cells(i, 9).Value
            End If
            
    Next i
    
    ws.Cells(2, 16).Value = ticker_max
    ws.Cells(3, 16).Value = ticker_min
    ws.Cells(4, 16).Value = ticker_total
    
    ws.Cells(2, 17).Value = increase_val
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).Value = decrease_val
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 17).Value = max_val

Next ws

End Sub



