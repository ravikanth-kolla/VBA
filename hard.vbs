Sub HardWorkBookCalculator()
        Dim ws As Worksheet
        For Each ws In ActiveWorkbook.Worksheets
            ws.Select
            Call Hard_Stock_Vol_Calculator_with_pct_change_and_greatest_vals
        Next ws
End Sub

Sub Hard_Stock_Vol_Calculator_with_pct_change_and_greatest_vals()
    Dim cur_stock_sym As String
    Dim tot_stock_vol As Double
    Dim ticker_count As Long
    Dim year_open_price As Double
    Dim year_close_price As Double
    Dim yearly_pct_change As Double
    
    Dim greatest_vol As Double
    Dim greatest_vol_stock As String
    
    Dim greatest_pct_inc As Double
    Dim greatest_pct_inc_stock As String
    
    Dim greatest_pct_dec As Double
    Dim greatest_pct_dec_stock As String
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percentage Change"
    
    Cells(1, 12).Value = "Total Volume"
    
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
            
             
            'Now check the greatest and least of all values
            If (tot_stock_vol > greatest_vol) Then
                greatest_vol = tot_stock_vol
                greatest_vol_stock = cur_stock_sym
            End If
            
            If (year_open_price <> 0) Then
                If (yearly_change / year_open_price > greatest_pct_inc) Then
                    greatest_pct_inc = yearly_change / year_open_price
                    greatest_pct_inc_stock = cur_stock_sym
                End If
            End If
            
            If (year_open_price <> 0) Then
                If (yearly_change / year_open_price < greatest_pct_dec) Then
                    greatest_pct_dec = yearly_change / year_open_price
                    greatest_pct_dec_stock = cur_stock_sym
                End If
            End If
            
            'Reset values
            tot_stock_vol = 0
            year_open_price = next_open_price

            
        End If
        
        
    Next nRow
    
    'Populate the extreme values
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    Cells(2, 17).Value = greatest_pct_inc
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(2, 16).Value = greatest_pct_inc_stock
    
    Cells(3, 17).Value = greatest_pct_dec
    Cells(3, 17).NumberFormat = "0.00%"
    Cells(3, 16).Value = greatest_pct_dec_stock
    
    Cells(4, 17).Value = greatest_vol
    Cells(4, 16).Value = greatest_vol_stock
    
    
End Sub

