Attribute VB_Name = "Module1"
Sub StockMarketAnalysis():
 
 
    Dim Openprice As Double
    Openprice = 0
    Dim closeprice As Double
    close_price = 0
    Dim YearlyChange As Double
    YearlyChange = 0
    Dim PercentChange As Double
    PercentChange = 0
    
    Dim Ticker As String
    Ticker = 0
    Dim TotalStockVol As Double
    TotalStockVol = 0
    Dim Value As Double
    Value = 0
    Dim StockSummaryTable As Long
    StockSummaryTable = 2

    Dim LRow As Long
    Dim LRowVal As Long

    Dim GreatestInc As Double
        GreatestInc = 0
    Dim GreatestDec As Double
        GreatestDec = 0
    Dim GreatestTotVol As Double
        GreatestTotVol = 0
    Dim TickerRow As Long:
        TickerRow = 1
'----------------------------------------

    For Each ws In Worksheets

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent change"
    ws.Range("L1").Value = "Total Stock Volume"

    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest total volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

    LRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    CounterTicker = 2
    CounterVolume = 2
    CounterYearly = 2
    
    For i = 2 To LRow
 
    If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        TickerRow = TickerRow + 1
        Openprice = ws.Cells(i, 3).Value
        CounterTicker = ws.Cells(i, 1).Value
        ws.Cells(TickerRow, "I").Value = Ticker
            End If
            
    If ws.Cells(i + 1).Value = ws.Cells(i, 1).Value Then
        volume = volume + ws.Cells(i, 7).Value
    Else
        volume = volume + ws.Cells(i, 7).Value
        ws.Cells(CounterVolume, 12).Value = Format(volume, "#.##0")
        If volume > GTotalVolume Then
            GTotalVolume = volume
            GTotalTicker = ws.Cells(i, 1).Value
        End If
        CounterVolume = CounterVolume + 1
        volume = 0
End If

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        closeprice = ws.Cells(2, 6).Value
        YearlyChange = closeprice - Openprice
        If Openprice <> 0 Then
            PercentageChange = (YearlyChange / Openprice) * 100
            If PercentageChange > Gincrease Then
            Gincrease = PercentageChange
            GincreaseTicker = ws.Cells(i, 1).Value
        ElseIf PercentageChange < GDecrease Then
        GDecrease = PercentageChange
        GDecreaseTicker = ws.Cells(i, 1).Value
        End If
    Else
        PercentageChange = 0
    End If
    ws.Cells(CounterYearly, 10).Value = Format(YearlyChange, "#.00")
    ws.Cells(CounterYearly, 11).Value = Format(PercentageChange, " 0.00") & "%"
    If ws.Cells(CounterYearly, 10).Value < 0 Then
        ws.Cells(CounterYearly, 10).Interior.ColorIndex = 3
        ws.Cells(CounterYearly, 11).Interior.ColorIndex = 3
        
    Else
        ws.Cells(CounterYearly, 10).Interior.ColorIndex = 4
        ws.Cells(CounterYearly, 11).Interior.ColorIndex = 4
    End If
    
        Openprice = 0
        closeprice = 0
        YearlyChange = 0
        PercentageChange = 0
        CounterYearly = CounterYearly + 1
    
      End If
    
        Next i
    Next ws

    Range("F2").Value = GIncreaseTikcer
    Range("F3").Value = GDecreaseTicker
    Range("F4").Value = GTotalTicker
    Range("Q2").Value = Format(Gincrease, "0.00") & "%"
    Range("F3").Value = Format(GDecrease, "0.00") & "%"
    Range("Q4").Value = Format(GTotalVolume, "#, ##0")
    
End Sub
