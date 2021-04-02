Sub StocksChallenge()

'for each loop to iterate through worksheets
  For Each WS In Worksheets
  
'declare variables
    Dim yearlyOpenPrice, yearlyClosingPrice, yearlyChange, yearlyPercent As Double
    Dim totalStockVolume, resultsRow As Long

'
    resultsRow = 2
    WS.Cells(1, 9).Value = "Ticker"
    WS.Cells(1, 10).Value = "Yearly Change"
    WS.Cells(1, 11).Value = "Percent Change"
    WS.Cells(1, 12).Value = "Total Stock Volume"
    
    For Row = 2 To (WS.Cells(Rows.Count, 1).End(xlUp).Row)
       
      If WS.Cells(Row, 1).Value <> WS.Cells(Row - 1, 1).Value Then
        totalStockVolume = WS.Cells(Row, 7).Value
        yearlyOpenPrice = WS.Cells(Row, 3).Value
        WS.Cells(resultsRow, 9).Value = WS.Cells(Row, 1).Value
        
      ElseIf WS.Cells(Row, 1).Value <> WS.Cells(Row + 1, 1).Value Then
        totalStockVolume = totalStockVolume + WS.Cells(Row, 7).Value
        WS.Cells(resultsRow, 12) = totalStockVolume
        
        yearlyClosingPrice = WS.Cells(Row, 6).Value
        
        yearlyChange = yearlyClosingPrice - yearlyOpenPrice
        WS.Cells(resultsRow, 10) = yearlyChange
        If yearlyChange < 0 Then
          WS.Cells(resultsRow, 10).Interior.ColorIndex = 3
          WS.Cells(resultsRow, 10).Font.ColorIndex = 1
        Else
          WS.Cells(resultsRow, 10).Interior.ColorIndex = 4
          WS.Cells(resultsRow, 10).Font.ColorIndex = 1
        End If
        If yearlyOpenPrice = 0 Then
            yearlyPercent = 0
        Else
            yearlyPercent = yearlyChange / yearlyOpenPrice
        End If
          
        WS.Cells(resultsRow, 11).Value = yearlyPercent
        WS.Cells(resultsRow, 11) = FormatPercent(WS.Cells(resultsRow, 11), 2)
      
        totalStockVolume = 0
        resultsRow = resultsRow + 1
        
      Else
        totalStockVolume = totalStockVolume + WS.Cells(Row, 7).Value
      
      End If
    Next Row
    
  Next WS
End Sub
