Sub StockData()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Yearly_change As Double
    Dim Percent_change As Double
    Dim Total_stock_volume As Double
    Dim Setter As Long
    Dim Open_price As Double
    Dim Close_price As Double
    Dim End_count As Long
    Dim i As Long
    Dim o As Long
    Dim j As Long
    Dim k As Long
    Dim Total_count As Long
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_total As Double
    Dim increase_ticker As String
    Dim decrease_ticker As String
    Dim greatest_ticker As String

    For Each ws In Worksheets

        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"

        Total_stock_volume = 0
        Setter = 2
        Open_price = ws.Cells(2, 3) ' Assuming row 2 is the first data row
        Close_price = 0
        End_count = 0
        
        For i = 2 To LastRow ' Assuming row 2 is the first data row
            Total_stock_volume = Total_stock_volume + ws.Cells(i, 7).Value
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
                Close_price = ws.Cells(i, 6)
                ws.Cells(Setter, 9) = ws.Cells(i, 1)
                
                Yearly_change = Close_price - Open_price
                ws.Cells(Setter, 10) = Yearly_change
                
                If Open_price = 0 Then
                    ws.Cells(Setter, 11) = 0
                Else
                    Percent_change = (Yearly_change / Open_price) * 100
                    ws.Cells(Setter, 11) = Percent_change
                End If

                ws.Cells(Setter, 12) = Total_stock_volume
                
                ' Reset for the next ticker
                Total_stock_volume = 0
                If i + 1 <= LastRow Then
                    Open_price = ws.Cells(i + 1, 3)
                End If
                Setter = Setter + 1
            End If
        Next i
        
        For o = 2 To Setter-1
            If ws.Cells(o, 10) > 0 Then
                ws.Cells(o, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(o, 10) < 0 Then
                ws.Cells(o, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(o, 10).Interior.ColorIndex = 0
            End If
        Next o
            
        greatest_increase = 0
        greatest_decrease = 0
        greatest_total = 0
        
        ws.Cells(1, 16) = "Ticker"
        ws.Cells(1, 17) = "Value"
        ws.Cells(2, 15) = "Greatest % increase"
        ws.Cells(3, 15) = "Greatest % decrease"
        ws.Cells(4, 15) = "Greatest Total Volume"
        
        For k = 2 To Setter-1
            If ws.Cells(k, 11) > greatest_increase Then
                greatest_increase = ws.Cells(k, 11)
                increase_ticker = ws.Cells(k, 9)
            End If
            
            If ws.Cells(k, 11) < greatest_decrease Then
                greatest_decrease = ws.Cells(k, 11)
                decrease_ticker = ws.Cells(k, 9)
            End If
            
            If ws.Cells(k, 12) > greatest_total Then
                greatest_total = ws.Cells(k, 12)
                greatest_ticker = ws.Cells(k, 9)
            End If
        Next k

        ws.Cells(2, 17) = greatest_increase
        ws.Cells(2, 16) = increase_ticker
        ws.Cells(3, 17) = greatest_decrease
        ws.Cells(3, 16) = decrease_ticker
        ws.Cells(4, 17) = greatest_total
        ws.Cells(4, 16) = greatest_ticker
    
    Next ws

End Sub
