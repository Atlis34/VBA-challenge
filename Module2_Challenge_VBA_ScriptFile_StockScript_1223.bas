Attribute VB_Name = "Module1"
Sub StockScript()

    For Each WS In Worksheets
        WS.Activate
            
      ' Set variables
      Dim SummaryRow As Integer
      Dim MaxRow As Long
      Dim OpenPrice As Double
      Dim ClosePrice As Double
      Dim YearlyChange As Double
      Dim PercentChange As Double
      Dim GreatestIncrease As Double
      Dim GreatestDecrease As Double
            
        'Set Variables to initial values
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestVolume = 0
        Totalvol = 0
        SummaryRow = 2
        row_count = Cells(Rows.Count, "A").End(xlUp).Row
        
      'Set Header row for summary section
      Cells(1, "I").Value = "Ticker"
      Cells(1, "J").Value = "Yearly Change"
      Cells(1, "K").Value = "Percent Change"
      Cells(1, "L").Value = "Total Stock Volume"
      
      'Set Cells for 'Greatest' section
      Cells(2, "O").Value = "Greatest % Increase"
      Cells(3, "O").Value = "Greatest % Decrease"
      Cells(4, "O").Value = "Greatest Total Volume"
      Cells(1, "P").Value = "Ticker"
      Cells(1, "Q").Value = "Value"
      
      'Set OpenPrice on first occurence of Ticker
      OpenPrice = Cells(2, "C").Value
        
      ' Loop through rows in the column
      
      RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        
        'Loop through rows in the column
        For i = 2 To RowCount
            Totalvol = Totalvol + Cells(i, "G").Value
        
          'When the company changes, then write the total out in column g
          If Cells(i + 1, "A").Value <> Cells(i, "A").Value Then
          'Get ClosePrice
              ClosePrice = Cells(i, "F").Value
              YearlyChange = ClosePrice - OpenPrice
              PercentChange = YearlyChange / OpenPrice * 100
              
              Cells(SummaryRow, "I").Value = Cells(i, "A").Value
              Cells(SummaryRow, "J").Value = YearlyChange
              Cells(SummaryRow, "K").Value = "%" & PercentChange
              Cells(SummaryRow, "L").Value = Totalvol
              
              'Assign green and red
              If YearlyChange > 0 Then
                    Cells(SummaryRow, "J").Interior.ColorIndex = 4
                ElseIf YearlyChange < 0 Then
                    Cells(SummaryRow, "J").Interior.ColorIndex = 3
                Else
                    Cells(SummaryRow, "J").Interior.ColorIndex = 2
              End If
              
              SummaryRow = SummaryRow + 1
              Totalvol = 0
              OpenPrice = Cells(i + 1, "C").Value
                        
    
          End If
    
        Next i
        
        'Set RowCount back to 0
         RowCount = 0
         'Set RowCount to max of Summary Table
         RowCount = Cells(Rows.Count, "I").End(xlUp).Row
         
        
        'Check for highest increase and decrease and assign values to greatest section for percentages
        For i = 2 To RowCount
            If Cells(i, "K").Value > 0 And Cells(i, "K").Value > GreatestIncrease Then
               GreatestIncrease = Cells(i, "K").Value
               Cells(2, "P").Value = Cells(i, "I").Value
             ElseIf Cells(i, "K").Value < 0 And Cells(i, "K").Value < GreatestDecrease Then
               GreatestDecrease = Cells(i, "K").Value
               Cells(3, "P").Value = Cells(i, "I").Value
            End If
            
            If Cells(i, "L").Value > GreatestVolume Then
                GreatestVolume = Cells(i, "L").Value
                Cells(4, "P").Value = Cells(i, "I").Value
            End If
            
        Next i
        
        'Assign Greatest Cell Values
        Cells(2, "Q").Value = "%" & GreatestIncrease * 100
        Cells(3, "Q").Value = "%" & GreatestDecrease * 100
        Cells(4, "Q").Value = GreatestVolume
        
      'Autofit columns
      Columns("A:Q").AutoFit

    
    Next WS

    MsgBox ("Complete")

End Sub

