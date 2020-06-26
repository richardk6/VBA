Attribute VB_Name = "Module1"
Sub ProcessStocks()

    ' Define ws
    Dim ws As Worksheet
    
    ' Define LastRow
    Dim LastRow As Double
    
        For Each ws In Worksheets
    
            ws.Activate
    
                Debug.Print (ws.Name)
            
                For j = 2 To LastRow
    
                    Debug.Print (ws.Cells(j, 1).Value)
            
                Next j
        
    ' Define ticker symbol
        Dim i As Long

    ' Define StartValue
        Dim StartValue As Long
    
    ' Define opening value
        Dim openvalue As Double
        
    ' Define variable for YearOpen
        Dim YearlyChange As Double
        
    ' Define Perfect Change
        Dim PercentChange As Double
        
    ' Define variable for StockVolumn
        Dim StockVolumn As Double
    
    ' Define variable for holding the ticker symbol
        Dim Ticker_Symbol As String
    
     ' Define Greatest % Increase
        Dim GreatestIncrease As Double
        
     ' Define
        Dim LargestStockVolumn As Double
        
     ' Define Greatest % Decrease
       Dim GreatestDecrease As Double
    
     ' Track location of where ticker symbols should go into spreadsheet
      Dim Ticker_Table_Row As Integer
    
        Ticker_Table_Row = 0
        StartValue = 2

    ' Define a variable for column of interest
     Dim column As Integer
         column = 1
         
    ' Define OpenValue
    openvalue = Cells(StartValue, 3).Value
    
    ' Giving value to StockVolumn
    StockVolumn = 0
   
    ' Determine the Last Row
     LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Create a loop from 2 to last row (put last row equation in) in column add was
         For i = 2 To LastRow
    
            ' Review Ticker Column by Searches for when the value of the next cell is different from the previous cell
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
            ' Look for the stock volumn
            StockVolumn = StockVolumn + Cells(i, 7).Value
                
                ' Put values into column I
                ' Set the Ticker Symbol
                Ticker_Symbol = Cells(i, 1).Value
                
                    If openvalue = 0 Then
                        StartValue = Cells(StartValue + 1, 3).Value
                    
                    Else
                    ' Percent Change
                    PercentChange = YearlyChange / openvalue
                    
                    End If
                    
                    ' Format to %
                    Cells(i, 11) = Format(PercentChange, "Percent")

                ' Yearly Change
                YearlyChange = Cells(i, 6).Value - Cells(StartValue, 3)
                
                ' Place ticker symbols into ticker column
                Range("I" & 2 + Ticker_Table_Row).Value = Ticker_Symbol
                
                ' Place stock volumn in column
                Range("L" & 2 + Ticker_Table_Row).Value = StockVolumn
                
                ' Place YearlyChange
                Range("J" & 2 + Ticker_Table_Row).Value = YearlyChange
                
                ' Place Percent Change
                Range("K" & 2 + Ticker_Table_Row).Value = PercentChange
                
                ' Place GreatIncreast
                Cells(2, 16).Value = GreatestIncrease
                
                 ' Format to %
                    Cells(2, 16) = Format(PercentChange, "Percent")
                
                ' Place Greatest Decrease
                Cells(3, 16).Value = GreatestDecrease
                
                 ' Format to %
                    Cells(3, 16) = Format(PercentChange, "Percent")
                
                ' Place LargestStockVolumn
                Cells(4, 16).Value = LargestStockVolumn
                
                ' Continue down the ticker column
                Ticker_Table_Row = Ticker_Table_Row + 1
                
                ' Reset Startvalue
                StartValue = i + 1
                
                'Reset StockVolumn
                StockVolumn = 0
                
                Else
                StockVolume = StockVolume + Cells(i, 7).Value
    
            End If
            
        Next i
    
                'Find the Greatest Increase in PercentChange
                Cells(2, 16).Value = WorksheetFunction.Max(Range("K2:K" & LastRow))
                
                ' Find the matching Ticker symbol for Greatest Increase for PercentChange
                ' Cells(2, 15).Value = WorksheetFunction.Match(Range("P2").Value, Range("I2:I" & LastRow), 0)
                
                'Find the Greatest Decrease in PercentChange
                Cells(3, 16).Value = WorksheetFunction.Min(Range("K2:K" & LastRow))
                
                ' Find the matching Ticker symbol for Greatest Decrease for PercentChange
                
                ' Find Largest Stock Volumn
                Cells(4, 16).Value = WorksheetFunction.Max(Range("L2:L" & LastRow))
                
                ' Find the matching Ticker symbol for Largest Stock Volumn
                
                ' Column titles
                Cells(1, 9).Value = "Ticker"
                Cells(1, 10).Value = "Yearly Change"
                Cells(1, 11).Value = "Percent Change"
                Cells(1, 12).Value = "Stock Volumn"
                Cells(1, 15).Value = "Ticker Symbol"
                Cells(1, 16).Value = "Value"
                Cells(2, 14).Value = "Greatest Percent Increase"
                Cells(3, 14).Value = "Greatest Percent Decrease"
                Cells(4, 14).Value = "Greatest Stock Volumn"
        
         ' Change the color of YearlyChange column
         
         For i = 2 To LastRow
        
            If Cells(i, 10).Value < 0 Then
            
                Cells(i, 10).Interior.ColorIndex = 3
                        
                Else
                Cells(i, 10).Interior.ColorIndex = 4
               
            End If
            
         Next i
         
    Next ws
    
End Sub
    

