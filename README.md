# VBA-Challenge
Sub StockData()

For Each ws In Worksheets

     Dim WSName As String
        WSName = ws.Name
    Dim StockTick As String
    Dim YearlyChange As Double
        YearlyChange = 0
    Dim PriceOpen As Double
        PriceOpen = ws.Cells(2, 3).Value
    Dim PriceClose As Double
        PriceClose = 0
    Dim SummaryTable As Integer
        SummaryTable = 2
    Dim StockVolume As Double
        StockVolume = 0
        PercentChange = 0
    Dim PercentOpen As Double
        PercentOpen = ws.Cells(2, 3).Value
    Dim PercentClose As Double
        PercentClose = 0
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    ws.Cells(1, 10) = "Ticker"
    ws.Cells(1, 11) = "Yearly Change"
    ws.Cells(1, 12) = "Percent Change"
    ws.Cells(1, 13) = "Total Stock Volume"
    
    
    Dim GreatestInc As Double
        GreatestInc = 0
    Dim TickerInc As String
    Dim GreatestDec As Double
        GreatestDec = 0
    Dim TickerDec As String
    Dim GreatTot As Double
        GreatTot = 0
    Dim TickerTotal As String
        
    For i = 2 To LastRow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            StockTick = ws.Cells(i, 1).Value
            ws.Range("J" & SummaryTable).Value = StockTick
            
            PriceClose = ws.Cells(i, 6).Value
            YearlyChange = PriceClose - PriceOpen
            ws.Range("K" & SummaryTable) = YearlyChange
            
                If YearlyChange > 0 Then
                    ws.Range("K" & SummaryTable).Interior.ColorIndex = 4
                Else
                    ws.Range("K" & SummaryTable).Interior.ColorIndex = 3
                End If
                
            
            
            
        
            PercentChange = YearlyChange / PriceOpen
            
                If PercentChange > GreatestInc Then
                    GreatestInc = PercentChange
                    TickerInc = ws.Cells(i, 1).Value
                Else
                    GreatestInc = GreatestInc
                    TickerInc = TickerInc
                End If
                 
                
                If GreatestDec > PercentChange Then
                    GreatestDec = PercentChange
                    TickerDec = ws.Cells(i, 1).Value
                Else
                    GreatestDec = GreatestDec
                    TickerDec = TickerDec
                End If
                
                PercentChange = Format(PercentChange, "0.0000%")
                ws.Range("L" & SummaryTable) = PercentChange
                PercentChange = 0
                PercentChange = ws.Cells(i + 1, 3)
                
                StockVolume = StockVolume + ws.Cells(i, 7).Value
                
                    If StockVolume > GreatTot Then
                        GreatTot = StockVolume
                        TickerTotal = ws.Cells(i, 1).Value
                    Else
                        GreatTot = GreatTot
                        TickerTotal = TickerTotal
                    End If
                ws.Range("M" & SummaryTable) = StockVolume
                StockVolume = 0
                
                PriceOpen = ws.Cells(i + 1, 3)
                SummaryTable = SummaryTable + 1
                
                
            Else
            
                StockVolume = StockVolume + Cells(i, 7).Value
                
            End If
            
        Next i
        
        ws.Cells(2, 16) = "Greatest % Increase"
        ws.Cells(3, 16) = "Greatest % Decrease"
        ws.Cells(4, 16) = "Greatest Total Volume"
        ws.Cells(1, 17) = "Ticker"
        ws.Cells(1, 18) = "Value"
        
        ws.Cells(2, 17).Value = TickerInc
        ws.Cells(3, 17).Value = TickerDec
        ws.Cells(4, 17).Value = TickerTotal
        ws.Cells(2, 18).Value = Format(GreatestInc, "0.00%")
        ws.Cells(3, 18).Value = Format(GreatestDec, "0.00%")
        ws.Cells(4, 18).Value = GreatTot
        
    Next ws
        



End Sub


[!Note]
I had trouble trying to write this code so I found a repository in GitHub by anniedonnelly which is what I used to help write my code. After I finished writing my code the output of my data on the excel sheet was not correct so I got some help from a learning assistant on AskBCS in slack where he helped me fix my problems with my percent change collumn and my bonus table. 

