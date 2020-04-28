Sub Vba_Stocks():


For Each ws In Worksheets
  
    'headers in stocks table
    ws.Cells(1, 9).Value = "Ticker Symbol"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
        
    'variables for stock table
    Dim ticker_symbol As String
    Dim yearly_change As Double
    Dim percent_change As Double
            
    Dim row_results As Integer
    row_results = 2
            
    Dim total_vol As Double
    total_vol = 0
            
    'variables for open price and close price
    Dim open_price As Double
    Dim close_price As Double
          
    'tickers to keep track of row number
    Dim ticker_number As Double
    ticker_number = 0
            
                                                            
    'Loop through rows
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
    For i = 2 To LastRow
               
        If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                           
            'ticker symbol and print row results
            ticker_symbol = ws.Cells(i, 1).Value
            ws.Range("I" & row_results).Value = ticker_symbol
                  
            'open price and close price
            open_price = ws.Cells(i - ticker_number, 3).Value
            close_price = ws.Cells(i, 6).Value
                    
            'yearly change and print in row results
            yearly_change = close_price - open_price                   
            ws.Range("J" & row_results).Value = yearly_change
                           
                'Color Yearly Change Column
                If yearly_change >= 0 Then
                    ws.Range("J" & row_results).Interior.Color = vbGreen
                Else
                    ws.Range("J" & row_results).Interior.Color = vbRed
                End If
                                   
                           
                           
            'percent change and print row results
            If (yearly_change = 0) Or (open_price = 0) Then
                ws.Range("K" & row_results).Value = 0
                        
            Else
                percent_change = yearly_change / open_price                           
                ws.Range("K" & row_results).Value = percent_change
                ws.Range("K" & row_results).NumberFormat = "0.00%"
            End If
                    
            'total stock volume in row results.
            total_vol = total_vol + ws.Cells(i, 7).Value                        
            ws.Range("L" & row_results).Value = total_vol
                       
        
                ' row results summary
                row_results = row_results + 1                            
                    
                'reset total stock
                total_vol = 0
                        
                'reset ticker_number
                ticker_number = 0
             
        Else
            total_vol = total_vol + ws.Cells(i, 7).Value
            ticker_number = ticker_vol + 1
                   
        End If
            
    Next i
    Next ws

End sub 
        

