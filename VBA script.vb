Sub StockHW():
    For Each ws In Worksheets
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        OpeningPrice = ws.Cells(2, 3).Value
        TotalVolume = 0
        Row = 2
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                ws.Cells(Row, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(Row, 10).Value = ws.Cells(i, 6).Value - OpeningPrice
                If ws.Cells(Row, 10).Value > 0 Then
                    ws.Cells(Row, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(Row, 10).Value < 0 Then
                    ws.Cells(Row, 10).Interior.ColorIndex = 3
                End If
                ws.Cells(Row, 10).NumberFormat = "0.00"
                ws.Cells(Row, 11).Value = ws.Cells(Row, 10).Value / OpeningPrice
                ws.Cells(Row, 11).Value = Format(ws.Cells(Row, 11).Value, "Percent")
                ws.Cells(Row, 12).Value = TotalVolume
                OpeningPrice = ws.Cells(i + 1, 3).Value
                TotalVolume = 0
                Row = Row + 1
            Else
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            End If
        Next i

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 17).Value = 0
        ws.Cells(3, 17).Value = 0
        ws.Cells(4, 17).Value = 0
        LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For p = 2 To LastRow2
            If ws.Cells(p + 1, 11).Value > ws.Cells(p, 11).Value And ws.Cells(p + 1, 11).Value > ws.Cells(2, 17).Value Then
                ws.Cells(2, 17).Value = ws.Cells(p + 1, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(p + 1, 9).Value
            End If
        Next p
        
        For m = 2 To LastRow2
            If ws.Cells(m + 1, 11).Value < ws.Cells(m, 11).Value And ws.Cells(m + 1, 11).Value < ws.Cells(3, 17).Value Then
                ws.Cells(3, 17).Value = ws.Cells(m + 1, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(m + 1, 9).Value
            End If
        Next m
        
        For n = 2 To LastRow2
            If ws.Cells(n + 1, 12).Value > ws.Cells(n, 12).Value And ws.Cells(n + 1, 12).Value > ws.Cells(4, 17).Value Then
                ws.Cells(4, 17).Value = ws.Cells(n + 1, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(n + 1, 9).Value
            End If
        Next n

        ws.Cells(2, 17).Value = Format(ws.Cells(2, 17).Value, "Percent")
        ws.Cells(3, 17).Value = Format(ws.Cells(3, 17).Value, "Percent")
        ws.Columns("A:Q").AutoFit
    
    Next ws

End Sub
