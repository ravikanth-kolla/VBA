Sub EasyWorkBookCalculator()
        Dim ws As Worksheet
        For Each ws In ActiveWorkbook.Worksheets
            ws.Select
            Call easy_tot_stock_vol
        Next ws
End Sub

Sub easy_tot_stock_vol()
    Dim cur_stock_sym As String
    Dim tot_stock_vol As Double
    Dim next_stock_sym As String
    Dim ticker_count As Long
            
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total Stock Volume"

    Dim lastrow As Double
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row

    For nRow = 2 To lastrow
        day_stock_vol = Cells(nRow, 7).Value
        cur_stock_sym = Cells(nRow, 1).Value
        next_stock_sym = Cells(nRow + 1, 1).Value
        tot_stock_vol = day_stock_vol + tot_stock_vol
        If (cur_stock_sym <> next_stock_sym) Then
            Cells(ticker_count + 2, 9).Value = cur_stock_sym
            Cells(ticker_count + 2, 10).Value = tot_stock_vol
            ticker_count = ticker_count + 1
            
            'Reset total volume
            tot_stock_vol = 0
        End If
        
    Next nRow
End Sub