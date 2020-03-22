Sub vba_stocks()

        'Set variables
        Dim ticker_result As String
        
        'Define row_results
        Dim row_results As Double
        row_results = row_results + 1
        
        'Dim rowcount As Long
        'rowcount = 2
        
         'Year Open
        Dim year_open As Double
        year_open = 0

        'Year close
        Dim year_close As Double
        year_close = 0
        
        'Year change
        Dim year_change As Double
        year_change = 0

        Dim total_vol As Double
        total_vol = 0

        'percent change in price for the year
        Dim percent_change As Double
        percent_change = 0

        'Set variable for total rows to loop through
        Dim lastrow As Long
        'lastrow = ws.Cells(Rows.Results, 1).End(xlUp).Row
        
    'Loop for worksheets
    'Dim ws As Worksheet
        'For Each ws In Worksheet

        'Create column headers for variables
        'ws.Cells(1, 9).Value = "Ticker"
        'ws.Cells(1, 10).Value = "Yearly Change"
        'ws.Cells(1, 11).Value = "Percent Change"
        'ws.Cells(1, 12).Value = "Total Stock Volume"

        'Loop to search through ticker symbols
        For i = 2 To lastrow
            
            'Conditional to grab year open price
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

                year_open = ws.Cells(i, 3).Value

            End If

            'total stock volume for the year
            total_vol = total_vol + ws.Cells(i, 7)

            'Conditional to determine if the ticker symbol is changing
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

                'Move ticker symbol to summary table
                ws.Cells(rowresults, 9).Value = ws.Cells(i, 1).Value

                'Move total stock volume to the summary table
                ws.Cells(rowresults, 12).Value = total_vol

                'Grab year end price
                year_close = ws.Cells(i, 6).Value

                'Calculate the price change for the year
                year_change = year_close - year_open
                ws.Cells(rowresults, 10).Value = year_change

                'highlight positive or negative change in green and red
                If year_change >= 0 Then
                    ws.Cells(rowresults, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(rowresults, 10).Interior.ColorIndex = 3
                End If

                If year_open = 0 Then
                    'Percent change
                    Dim percent_change_NA As String
                    percent_change_NA = "New Stock"
                    ws.Cells(RowCount, 11).Value = percent_change
                Else
                    percent_change = year_change / year_open
                    ws.Cells(RowCount, 11).Value = percent_change
                    ws.Cells(RowCount, 11).NumberFormat = "0.00%"
                End If

            
            Next i
            
    'Next ws
        
        End Sub
