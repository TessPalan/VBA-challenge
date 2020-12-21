Sub VBAStocks():
    
    For Each ws In Worksheets
        
        ' add headers to summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' define dims
        Dim Ticker  As String
        Dim YearlyChange As Double
        Dim PercentageChange As Double
        Dim TotalStockVolume As Variant
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        
        Ticker = ws.Cells(2, 1).Value
        
        TotalStockVolume = 0
        SummaryTableRow = 2
        OpenRow = 2
        
        ' find data's last row
        LastRow = ws.Cells(1, 1).End(xlDown).Row
        ' MsgBox (LastRow)
        
        ' loop for ticker and calculate yearly change, percentage change and total volume
        For Row = 2 To LastRow
            
            TotalStockVolume = TotalStockVolume + ws.Cells(Row, 7).Value
            
            If ws.Cells(Row, 1) <> ws.Cells(Row + 1, 1) Then
                
                'ticker
                ws.Cells(SummaryTableRow, 9).Value = Ticker
                
                'yearly change
                OpenPrice = ws.Cells(OpenRow, 3).Value
                ClosePrice = ws.Cells(Row, 6).Value
                YearlyChange = ClosePrice - OpenPrice
                ws.Cells(SummaryTableRow, 10).Value = YearlyChange
                
                If ws.Cells(SummaryTableRow, 10).Value < 0 Then
                    
                    ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
                    
                ElseIf ws.Cells(SummaryTableRow, 10).Value > 0 Then
                    ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
                    
                End If
                
                'percentage change
                If OpenPrice <> 0 Then
                    PercentageChange = YearlyChange / OpenPrice
                Else
                    PercentageChange = 0
                End If
                
                ws.Cells(SummaryTableRow, 11).Value = PercentageChange
                ws.Cells(SummaryTableRow, 11) = Format(PercentageChange, "Percent")
                
                ' total volume
                ws.Cells(SummaryTableRow, 12).Value = TotalStockVolume
                TotalStockVolume = 0
                
                ' set back for next ticker calculation
                SummaryTableRow = SummaryTableRow + 1
                OpenRow = Row + 1
                
                Ticker = ws.Cells(Row + 1, 1).Value
            End If
            
        Next Row
        
    Next ws
    
End Sub