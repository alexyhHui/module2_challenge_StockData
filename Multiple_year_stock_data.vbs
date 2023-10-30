Sub StockAnalysis()
    
 For Each ws In Worksheets
    
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentageChange As Double
    Dim TotalVolume As Double
    
    OpeningPrice = 0
    ClosingPrice = 0
    YearlyChange = 0
    PercentageChange = 0
    TotalVolume = 0
    
    Dim MaxIncreaseTicker As String
    Dim MaxDecreaseTicker As String
    Dim MaxVolumeTicker As String
    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxVolume As Double
    
    MaxIncrease = 0
    MaxDecrease = 0
    MaxVolume = 0
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
   
    lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
   
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    OpeningPrice = ws.Cells(2, 3).Value
   
    For i = 2 To lastrow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      Ticker = ws.Cells(i, 1).Value
      ClosingPrice = ws.Cells(i, 6).Value

    YearlyChange = ClosingPrice - OpeningPrice
    
    If OpeningPrice <> 0 Then
       PercentageChange = (YearlyChange / OpeningPrice) * 100
    End If
      
      TotalVolume = TotalVolume + ws.Cells(i, 7).Value
      
      If YearlyChange > 0 Then
         ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
      ElseIf YearlyChange <= 0 Then
         ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
      End If
       
      ws.Range("I" & Summary_Table_Row).Value = Ticker
      ws.Range("L" & Summary_Table_Row).Value = TotalVolume
      ws.Range("J" & Summary_Table_Row).Value = YearlyChange
      ws.Range("K" & Summary_Table_Row).Value = PercentageChange & "%"
      
      If PercentageChange > MaxIncrease Then
          MaxIncrease = PercentageChange
          MaxIncreaseTicker = Ticker
        End If
                
        If PercentageChange < MaxDecrease Then
           MaxDecrease = PercentageChange
           MaxDecreaseTicker = Ticker
        End If
                
        If TotalVolume > MaxVolume Then
           MaxVolume = TotalVolume
           MaxVolumeTicker = Ticker
        End If
        
        ws.Range("P2").Value = MaxIncreaseTicker
        ws.Range("P3").Value = MaxDecreaseTicker
        ws.Range("P4").Value = MaxVolumeTicker
        ws.Range("Q2").Value = MaxIncrease & "%"
        ws.Range("Q3").Value = MaxDecrease & "%"
        ws.Range("Q4").Value = MaxVolume
      
      Summary_Table_Row = Summary_Table_Row + 1
      TotalVolume = 0
      YearlyChange = 0
      ClosingPrice = 0
      OpeningPrice = ws.Cells(i + 1, 3).Value
     
      Else

      TotalVolume = TotalVolume + ws.Cells(i, 7).Value
      
      End If

    Next i
  
  Next ws

End Sub