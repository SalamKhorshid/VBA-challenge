Sub Stock_Challegne()

'Loop through all sheets
For Each ws In Worksheets

'Set the first row of the summary table:
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"

'Set initial variables:
Dim ticker As String
Dim openPrice As Double
Dim closePrice As Double
Dim vol As LongLong
Dim yearlyChange As Double
Dim percentChange As Double
Dim totalVol As LongLong

'Keep track of the location for each ticker in the summary table:
Dim summary_table_row As Integer
summary_table_row = 2

'Set the last row:
Dim LR As LongLong
LR = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Loop through all tickers
        For i = 2 To LR
        
                'Check if we are still within the same ticker, if it is not ...
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                        'Set the ticker name in the summary table:
                        ticker = ws.Cells(i, 1).Value
                        
                        'Set a value for the open price:
                        If openPrice = 0 Then openPrice = ws.Cells(i, 3).Value
                        
                        'Find the close price in the last row (end of the year) of the same ticker:
                        If closePrice = 0 And ws.Cells(i + 1, 1).Value <> ticker Then
                            closePrice = ws.Cells(i, 6).Value
                            
                                'Calculate the Yearly Change:
                                yearlyChange = closePrice - openPrice
                                
                                'Print the total stock volume to the summary table:
                                 ws.Range("j" & summary_table_row).Value = yearlyChange
                        
                        'Calculate the percent change:
                        If openPrice <> 0 Then
                                percentChange = yearlyChange / openPrice
                        Else
                                percentChange = 0
                        End If
                        
                                'Conditional formatting to the percent change:
                                If percentChange >= 0 Then
                                    ws.Range("K" & summary_table_row).Interior.Color = vbGreen
                                ElseIf percentChange <= 0 Then
                                    ws.Range("K" & summary_table_row).Interior.Color = vbRed
                                Else
                                    'Leave the cell color as it is
                                    
                                End If
                        
                                'Print the ticker name to the summary table:
                                ws.Range("I" & summary_table_row).Value = ticker
                                
                                'Print the percent change to the summary table:
                                ws.Range("K" & summary_table_row).Value = FormatPercent(percentChange)
                        
                        'Calculate the total stock volume:
                        vol = vol + ws.Cells(i, 7).Value
                        
                                 'Print the total stock volume to the summary table:
                                 ws.Range("L" & summary_table_row).Value = vol
                                            
                        'Reset the open and close prices and total stock volume for the next ticker:
                        openPrice = 0
                        closePrice = 0
                        vol = 0
                        
                                 'Add one to the summary table row:
                                 summary_table_row = summary_table_row + 1
                                                     
                End If
                
        'If we are still within the same ticker
        Else
                
                        'Calculate the total stock volume:
                        vol = vol + ws.Cells(i, 7).Value
        
        End If
                
    Next i
    
                'Set the information of the statistical table:
                ws.Range("P1") = "Ticker"
                ws.Range("Q1") = "Value"
                ws.Range("O2") = "Greatest % Increase"
                ws.Range("O3") = "Greatest % Decrease"
                ws.Range("O4") = "Greatest Total Volume"
                
                'Set variables for the statistical table:
                Dim maxIncrease As Double
                Dim maxDecrease As Double
                Dim maxVolume As LongLong
                Dim tickerMaxIncrease As String
                Dim tkckerMaxDecrease As String
                Dim tickerMaxVolume As String
                
                'Initialize variables
                maxIncrease = 0
                maxDecrease = 0
                maxVolume = 0
                tickerMaxIncrease = ""
                tikckerMaxDecrease = ""
                tickerMaxVolum = ""
            
            'Set the last row of the summary table:
            Dim NLR As LongLong
            NLR = ws.Cells(Rows.Count, 9).End(xlUp).Row
            
            'Loop through all tickers in the summary table:
            For j = 2 To NLR
            
            'Set the ticker name in the summary table:
                        ticker = ws.Cells(j, 9).Value
            
            'Total volume for each ticker:
            totalVolume = ws.Cells(j, 12).Value
            
                    'Check if total Volume is greater than the previous maximum:
                    If totalVolume > maxVolume Then
                            maxVolume = totalVolume
                            tickerMaxVolume = ticker
                    End If
                    
            'Set the percent change in the summary table for each ticker:
            percentChangeST = ws.Cells(j, 11).Value
            
                    'Check if the percent change is greater than the prevouse maximum increase:
                    If percentChangeST > maxIncrease Then
                            maxIncrease = percentChangeST
                            tickerMaxIncrease = ticker
                    End If
                    
                    'Check if the percent change is less than the prevouse maximum decrease:
                    If percentChangeST < maxDecreasee Then
                            maxDecrease = percentChangeST
                            tikckerMaxDecrease = ticker
                    End If
        
            Next j
            
            'Print the findings of the statistical Table:
            ws.Range("P2").Value = tickerMaxIncrease
            ws.Range("P3").Value = tikckerMaxDecrease
            ws.Range("P4").Value = tickerMaxVolume
            ws.Range("Q2").Value = FormatPercent(maxIncrease)
            ws.Range("Q3").Value = FormatPercent(maxDecrease)
            ws.Range("Q4").Value = maxVolume
            
Next ws

End Sub
