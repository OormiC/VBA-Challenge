Attribute VB_Name = "Module1"
Sub practice_ticker()

Dim numrows As Long
Dim ticker_counter As Integer
Dim open_price As Double
Dim close_price As Double
Dim max_ticker As String
Dim min_ticker As String
Dim max_val_ticker As String
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets:
    'initialise variables
    numrows = Range("A1", Range("A1").End(xlDown)).Rows.Count
    ticker_counter = 2
    total_vol = 0
    open_price = ws.Cells(2, 3).Value
    
    'set the columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total stock volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % increase"
    ws.Range("O3").Value = "Greatest % decrease"
    ws.Range("O4").Value = "Max volume"
    
    
    
    For i = 2 To numrows:
        'add to the total volume for current ticker
        total_vol = total_vol + ws.Cells(i, 7).Value
    
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            'add ticker name and total stack volume
            ws.Cells(ticker_counter, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(ticker_counter, 12).Value = total_vol
        
            'calculate yearly and percent change
            close_price = ws.Cells(i, 6).Value
            ws.Cells(ticker_counter, 10).Value = close_price - open_price
            ws.Cells(ticker_counter, 11).Value = (close_price - open_price) / close_price
            ws.Cells(ticker_counter, 11).NumberFormat = "0.00%"
        
            'format colours of yearly and percent change
            If ws.Cells(ticker_counter, 10).Value < 0 Then
                ws.Cells(ticker_counter, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(ticker_counter, 10).Interior.ColorIndex = 4
            End If
        
            If Cells(ticker_counter, 11).Value < 0 Then
                ws.Cells(ticker_counter, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(ticker_counter, 11).Interior.ColorIndex = 4
            End If
        
            'reset variables for next ticker
            total_vol = 0
            ticker_counter = ticker_counter + 1
            open_price = ws.Cells(i + 1, 3).Value
        End If
        
    Next i
    
    'loop to find the greatest % increase
    Max = ws.Range("K2").Value
    For k = 2 To Range("K1", Range("K1").End(xlDown)).Rows.Count:
        If ws.Cells(k, 11).Value > Max Then
            Max = ws.Cells(k, 11).Value
            max_ticker = ws.Cells(k, 9).Value
        End If
    Next k
    
    'find greatest % decrease
    Min = ws.Range("K2").Value
    For l = 2 To Range("K1", Range("K1").End(xlDown)).Rows.Count:
        If ws.Cells(l, 11).Value < Min Then
            Min = ws.Cells(l, 11).Value
            min_ticker = ws.Cells(l, 9).Value
        End If
    Next l
    'find greatest total volume
    max_val = ws.Range("L2").Value
    For j = 2 To Range("L1", Range("L1").End(xlDown)).Rows.Count:
        If ws.Cells(j, 12).Value > max_val Then
            max_val = ws.Cells(j, 12).Value
            max_val_ticker = ws.Cells(j, 9).Value
        End If
    Next j
    
    ws.Range("Q2").Value = Max
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("P2").Value = max_ticker
    ws.Range("Q3").Value = Min
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("P3").Value = min_ticker
    ws.Range("Q4").Value = max_val
    ws.Range("P4").Value = max_val_ticker
Next ws
End Sub


