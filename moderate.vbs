Sub ModerateWorkBookCalculator()
        Dim ws As Worksheet
        For Each ws In ActiveWorkbook.Worksheets
            ws.Select
            Call Moderate_Stock_Vol_Calculator_with_pct_change
        Next ws
End Sub


Sub Moderate_Stock_Vol_Calculator_with_pct_change()
    Dim cur_stock_sym As String
    Dim tot_stock_vol As Double
    Dim prev_stock_sym As String
    Dim ticker_count As Long
    Dim year_open_price As Double
    Dim year_close_price As Double
    Dim yearly_pct_change As Double
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percentage Change"
    
    Cells(1, 12).Value = "Total Stock Volume"
    
    Dim lastrow As Double
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row

    'Open price for the year for 1st ticker
    year_open_price = Cells(2, 3).Value
    
    For nRow = 2 To lastrow
        cur_stock_sym = Cells(nRow, 1).Value
        day_stock_vol = Cells(nRow, 7).Value
        day_open_price = Cells(nRow, 3).Value
        day_close_price = Cells(nRow, 6).Value
        
        next_stock_sym = Cells(nRow + 1, 1).Value
        next_open_price = Cells(nRow + 1, 3).Value
        
        tot_stock_vol = day_stock_vol + tot_stock_vol
        If (year_open_price = 0) Then
            year_open_price = day_open_price
        End If
        
        If (cur_stock_sym <> next_stock_sym) Then
            year_close_price = day_close_price
            Cells(ticker_count + 2, 9).Value = cur_stock_sym
            Cells(ticker_count + 2, 12).Value = tot_stock_vol
            ticker_count = ticker_count + 1
            
            'calculate yearly and pct change
            yearly_change = (year_close_price - year_open_price)
            Cells(ticker_count + 1, 10).Value = yearly_change
            If (year_close_price < year_open_price) Then
                Cells(ticker_count + 1, 10).Interior.Color = vbRed
            Else
                Cells(ticker_count + 1, 10).Interior.Color = vbGreen
            
            End If
            
            If (year_open_price = 0) Then
                'Check for divide by zero error
                Cells(ticker_count + 1, 11).Value = "NA"
            Else
                Cells(ticker_count + 1, 11).Value = yearly_change / year_open_price
            End If
            Cells(ticker_count + 1, 11).NumberFormat = "0.00%"
            
            'Reset values
            tot_stock_vol = 0
            year_open_price = next_open_price
             
        End If
                       
    Next nRow
End Sub

