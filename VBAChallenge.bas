Sub VBAChallenge()


   Dim CurrentWs  As Worksheet

   For Each CurrentWs In Worksheets
        Dim TickerName As String
    
        Dim TickerVolume As Double
        TickerVolume = 0

        Dim summaryTickerRow As Integer
        summaryTickerRow = 2
        
        Dim OPrice As Double
        OPrice = CurrentWs.Cells(2, 3).Value
        
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double

        CurrentWs.Cells(1, 9).Value = "Ticker"
        CurrentWs.Cells(1, 10).Value = "Yearly Change"
        CurrentWs.Cells(1, 11).Value = "Percent Change"
        CurrentWs.Cells(1, 12).Value = "Total Stock Volume"

        Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To Lastrow

            If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
        
              TickerName = CurrentWs.Cells(i, 1).Value

              TickerVolume = TickerVolume + CurrentWs.Cells(i, 7).Value

              CurrentWs.Range("I" & summaryTickerRow).Value = TickerName

              CurrentWs.Range("L" & summaryTickerRow).Value = TickerVolume

              ClosePrice = CurrentWs.Cells(i, 6).Value

               YearlyChange = (ClosePrice - OPrice)
              
              CurrentWs.Range("J" & summaryTickerRow).Value = YearlyChange

                If OPrice = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = YearlyChange / OPrice
                End If

              CurrentWs.Range("K" & summaryTickerRow).Value = PercentChange
              CurrentWs.Range("K" & summaryTickerRow).NumberFormat = "0.00%"
   
              summaryTickerRow = summaryTickerRow + 1

              TickerVolume = 0

              OPrice = CurrentWs.Cells(i + 1, 3)
            
            Else
              
              TickerVolume = TickerVolume + CurrentWs.Cells(i, 7).Value

            
            End If
        
        Next i

    LastRowSTable = CurrentWs.Cells(Rows.Count, 9).End(xlUp).Row
    
        For i = 2 To LastRowSTable
            If CurrentWs.Cells(i, 10).Value > 0 Then
                CurrentWs.Cells(i, 10).Interior.ColorIndex = 10
            Else
                CurrentWs.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i


        CurrentWs.Cells(2, 15).Value = "Greatest % Increase"
        CurrentWs.Cells(3, 15).Value = "Greatest % Decrease"
        CurrentWs.Cells(4, 15).Value = "Greatest Total Volume"
        CurrentWs.Cells(1, 16).Value = "Ticker"
        CurrentWs.Cells(1, 17).Value = "Value"

        For i = 2 To LastRowSTable
            If CurrentWs.Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & LastRowSTable)) Then
                CurrentWs.Cells(2, 16).Value = CurrentWs.Cells(i, 9).Value
                CurrentWs.Cells(2, 17).Value = CurrentWs.Cells(i, 11).Value
                CurrentWs.Cells(2, 17).NumberFormat = "0.00%"

            ElseIf CurrentWs.Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & LastRowSTable)) Then
                CurrentWs.Cells(3, 16).Value = CurrentWs.Cells(i, 9).Value
                CurrentWs.Cells(3, 17).Value = CurrentWs.Cells(i, 11).Value
                CurrentWs.Cells(3, 17).NumberFormat = "0.00%"
            
            ElseIf CurrentWs.Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & LastRowSTable)) Then
                CurrentWs.Cells(4, 16).Value = CurrentWs.Cells(i, 9).Value
                CurrentWs.Cells(4, 17).Value = CurrentWs.Cells(i, 12).Value
            
            End If
        
        Next i
    Next CurrentWs
End Sub


