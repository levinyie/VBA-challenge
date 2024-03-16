Sub StockData():


    For Each ws In Worksheets
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    
    Dim Yearly_change As Double
    Dim Percent_change As Double
    Dim Total_stock_volume As Double
    Total_stock_volume = 0
    Dim Setter As Double
    Setter = 2
    Dim Open_price As Double
    Open_price = 0
    Dim Close_price As Double
    Close_price = 0
    Dim End_count As Integer
    End_count = 0
    
    For i = 1 To LastRow
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
            ws.Cells(Setter, 9) = ws.Cells(i + 1, 1)
            Open_price = ws.Cells(i + 1, 3)
            End_count = WorksheetFunction.CountIf(ws.Range("A:A"), ws.Cells(Setter, 9))
            Close_price = ws.Cells(i + End_count, 6)
            ws.Cells(Setter, 10) = Close_price - Open_price
            If Open_price = 0 Then
                ws.Cells(Setter, 11) = 0
            Else
                ws.Cells(Setter, 11) = ((Close_price - Open_price) / Open_price) * 100
            End If
            
            Setter = Setter + 1
            
        End If
            
    Next i
        
    For o = 2 To LastRow
        If ws.Cells(o, 10) > 0 Then
            ws.Cells(o, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(o, 10) < 0 Then
            ws.Cells(o, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(o, 10).Interior.ColorIndex = 0
        End If
        
    Next o
        
    Dim Total_count As Integer
    Total_count = 2
        
    For j = 2 To LastRow
        Total_stock_volume = Total_stock_volume + ws.Cells(j, 7).Value
            If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then
                ws.Cells(Total_count, 12).Value = Total_stock_volume
                Total_count = Total_count + 1
                Total_stock_volume = 0
            End If
    Next j
    
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_total As Double
    greatest_increase = 0
    greatest_decrease = 0
    greatest_total = 0
    Dim increase_ticker As String
    Dim decrease_ticker As String
    Dim greatest_ticker As String
    
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
    ws.Cells(2, 15) = "Greatest % increase"
    ws.Cells(3, 15) = "Greatest % decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
    
    For k = 2 To LastRow
        If ws.Cells(k, 11) > greatest_increase Then
            greatest_increase = ws.Cells(k, 11)
            increase_ticker = ws.Cells(k, 9)
        End If
        ws.Cells(2, 17) = greatest_increase
        ws.Cells(2, 16) = increase_ticker
        
        If ws.Cells(k, 11) < greatest_decrease Then
            greatest_decrease = ws.Cells(k, 11)
            decrease_ticker = ws.Cells(k, 9)
        End If
        ws.Cells(3, 17) = greatest_decrease
        ws.Cells(3, 16) = decrease_ticker
        
        If ws.Cells(k, 12) > greatest_total Then
            greatest_total = ws.Cells(k, 12)
            greatest_ticker = ws.Cells(k, 9)
        End If
        ws.Cells(4, 17) = greatest_total
        ws.Cells(4, 16) = greatest_ticker
    
    Next k
    
    Next ws

End Sub
