Sub LoopFin()

    For Each ws In Worksheets
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"

        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        Dim TickerName As String

        Dim TickerTotal As Double
        TickerTotal = 0

        Dim TickerChange As Double
        TickerChange = 0

        Dim TickerPerChange As Double
        TickerPerChange = 0

        Dim TickerOpen As Double
        TickerOpen = 0
        Dim TickerClose As Double
        TickerClose = 0

        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        stock_start_row = 2
        
            For i = 2 To lastrow

                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                    TickerName = ws.Cells(i, 1).Value

                    TickerChange = ws.Cells(i, 6).Value - ws.Cells(stock_start_row, 3).Value
                    
                    TickerPerChange = (ws.Cells(i, 6).Value - ws.Cells(stock_start_row, 3).Value) / (ws.Cells(stock_start_row, 3).Value - 0.0000001)

                    TickerTotal = TickerTotal + ws.Cells(i, 7).Value

                    ws.Range("I" & Summary_Table_Row).Value = TickerName
                    ws.Range("J" & Summary_Table_Row).Value = TickerChange
                        
                                If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                                    
                                ElseIf ws.Range("J" & Summary_Table_Row).Value < 0 Then
                                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                                        
                                End If
                                
                    ws.Range("K" & Summary_Table_Row).Value = TickerPerChange
                    ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                    ws.Range("L" & Summary_Table_Row).Value = TickerTotal

                    Summary_Table_Row = Summary_Table_Row + 1

                    TickerTotal = 0

                stock_start_row = i + 1

                Else

                    TickerTotal = TickerTotal + ws.Cells(i, 7)
                
                End If

            Next i

        ws.Cells(2, 15).Value = "Greatest % Increase"
        Dim GreatestInc As Double
        GreatestInc = ws.Cells(2, 16).Value
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
                

        ws.Columns("I:Q").AutoFit
       
    Next ws

End Sub




