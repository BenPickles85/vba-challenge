Attribute VB_Name = "Module3"
Sub StockTicker()
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % increase"
    ws.Cells(3, 14).Value = "Greatest % decrease"
    ws.Cells(4, 14).Value = "Greatest total volume"
    
    Dim i As Long
    Dim j As Long
    Dim StockVol As Double
    Dim End_Row As Long
    Dim start_price As Double
    Dim end_price As Double
    Dim price_change As Currency
    Dim percent_change As Double
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_volume As Double
    Dim ws As Worksheet
    
    For Each ws In Worksheets
       
    End_Row = Cells(Rows.Count, 1).End(xlUp).Row
    StockVol = 0
    j = 2
    start_price = Cells(2, 3).Value
        
    For i = 2 To End_Row
    
    'Calculate stock volume and print tickers
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value And ws.Cells(i, 3).Value <> 0 Then
            ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
            StockVol = StockVol + ws.Cells(i, 7).Value
            ws.Cells(j, 12).Value = StockVol
            
            end_price = ws.Cells(i, 6).Value
                
            If start_price <> 0 Then
                price_change = end_price - start_price
                ws.Cells(j, 10).Value = price_change
                percent_change = price_change / start_price
                ws.Cells(j, 11).Value = percent_change
                ws.Cells(j, 11).NumberFormat = "0.00%"
            End If
         
     'Conditional Formatting
            If ws.Cells(j, 11).Value > 0 Then
                ws.Cells(j, 11).Interior.Color = RGB(0, 230, 70)
            ElseIf ws.Cells(j, 11).Value < 0 Then
                    ws.Cells(j, 11).Interior.Color = RGB(255, 0, 0)
            End If
    
    'Reset
        StockVol = 0
        start_price = ws.Cells(i + 1, 3).Value
        j = j + 1
    
    'Stock Counter
        ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            StockVol = StockVol + ws.Cells(i, 7).Value
            
        End If
    
    Next i

    greatest_increase = 0
    greatest_decrease = 0
    greatest_volume = 0
    
    'Calculate and populate Greatest increase
    For i = 2 To End_Row
        If ws.Cells(i + 1, 11).Value > ws.Cells(i, 11).Value And ws.Cells(i + 1, 11).Value > greatest_increase Then
            greatest_increase = ws.Cells(i + 1, 11).Value
            ws.Cells(2, 16).Value = greatest_increase
            ws.Cells(2, 16).NumberFormat = "0.00%"
            ws.Cells(2, 15).Value = ws.Cells(i + 1, 9).Value
        End If
        
    'Calculate and populate Greatest increase
        If ws.Cells(i + 1, 11).Value < ws.Cells(i, 11).Value And ws.Cells(i + 1, 11).Value < greatest_decrease Then
            greatest_decrease = ws.Cells(i + 1, 11).Value
            ws.Cells(3, 16).Value = greatest_decrease
            ws.Cells(3, 16).NumberFormat = "0.00%"
            ws.Cells(3, 15).Value = ws.Cells(i + 1, 9).Value
        End If
    
    'Calculate and populate total volume
        If ws.Cells(i + 1, 12).Value > ws.Cells(i, 12).Value And ws.Cells(i + 1, 12).Value > greatest_volume Then
            greatest_volume = ws.Cells(i + 1, 12).Value
            ws.Cells(4, 16).Value = greatest_volume
            ws.Cells(4, 15).Value = ws.Cells(i + 1, 9).Value
        End If
        
    Next i
        
    Next ws
        
End Sub

