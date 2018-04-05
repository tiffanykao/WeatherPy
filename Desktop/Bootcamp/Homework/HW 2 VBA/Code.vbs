Sub StockData()
     
    For Each ws In Worksheets
     
        Dim sum_volume As Variant
        Dim sum_open As Variant
        Dim sum_close As Variant
        Dim yearly_change As Variant
     
        Dim ticker As String
        Dim data_last_row As Variant
        Dim ticker_last_row As Variant
     
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
     
        data_last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row

        For i = 2 To data_last_row
            ticker = ws.Cells(i, 1)
            sum_volume = ws.Cells(i, 7) + sum_volume
            sum_open = ws.Cells(i, 3) + sum_open
            sum_close = ws.Cells(i, 6) + sum_close
            ticker_last_row = ws.Cells(Rows.Count, "I").End(xlUp).Row
                If ws.Cells(i + 1, 1) <> ticker Then
                    ws.Cells(ticker_last_row + 1, 9) = ticker
                    ws.Cells(ticker_last_row + 1, 10) = sum_close - sum_open
                    If sum_close - sum_open < 0 Then
                        ws.Cells(ticker_last_row + 1, 10).Interior.ColorIndex = 3
                    Else
                        ws.Cells(ticker_last_row + 1, 10).Interior.ColorIndex = 4
                    End If
                    
                    ' If sum_open = 0 then we need the % change = 0 b/c you can't divide by 0. See 2014 data
                    If sum_open <> 0 Then
                        ws.Cells(ticker_last_row + 1, 11) = (sum_close - sum_open) / sum_open
                        ws.Cells(ticker_last_row + 1, 11).NumberFormat = "0.00%"
                    Else
                        ws.Cells(ticker_last_row + 1, 11) = 0
                    Endif

                    ws.Cells(ticker_last_row + 1, 12) = sum_volume
                    
                    sum_volume = 0
                    sum_open = 0
                    sum_close = 0
                End If
        Next i
    
        'Reset ticker_last_row after all the info is there
        ticker_last_row = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    
        Dim largest_percent_increase As Variant
        Dim largest_percent_decrease As Variant
        Dim largest_volume As Variant
        Dim largest_percent_increase_ticker As String
        Dim largest_percent_decrease_ticker As String
        Dim largest_volume_ticker As String
    
        largest_percent_increase = 0
        largest_percent_decrease = 0
        largest_volume = 0
        
        
        'Find largest_volume
        largest_volume = ws.Range("L2")
        largest_volume_ticker = ws.Range("I2")
        
        For i = 2 To ticker_last_row
            If ws.Cells(i, 12) > largest_volume Then
                largest_volume = ws.Cells(i, 12)
                largest_volume_ticker = ws.Cells(i, 9)
            End If
        Next i
        ws.Range("Q4") = largest_volume
        ws.Range("P4") = largest_volume_ticker
    
        'Find largest_percent_increase
        largest_percent_increase = ws.Range("K2")
        largest_percent_increase_ticker = ws.Range("I2")
        
        For i = 2 To ticker_last_row
            If ws.Cells(i, 11) > largest_percent_increase Then
                largest_percent_increase = ws.Cells(i, 11)
                largest_percent_increase_ticker = ws.Cells(i, 9)
            End If
        Next i
        ws.Range("Q2") = largest_percent_increase
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("P2") = largest_percent_increase_ticker
        
        'Find largest_percent_decrease
        largest_percent_decrease = ws.Range("K2")
        largest_percent_decrease_ticker = ws.Range("I2")
        
        For i = 2 To ticker_last_row
            If ws.Cells(i, 11) < largest_percent_decrease Then
                largest_percent_decrease = ws.Cells(i, 11)
                largest_percent_decrease_ticker = ws.Cells(i, 9)
            End If
        Next i
        ws.Range("Q3") = largest_percent_decrease
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("P3") = largest_percent_decrease_ticker
            
    Next ws

End Sub