' This script iterates through all stock data on a quarterly basis and performs the following tasks:
' 1. Identifies the ticker symbol for each stock.
' 2. Calculates the quarterly change in price by subtracting the opening price of the quarter
'    from the closing price of the quarter.
' 3. Computes the percentage change in price over the quarter based on the opening price.
' 4. Aggregates the total stock volume for the quarter.
' The results are then output for each stock, providing a summary of its quarterly performance.

Sub QuarterlyStockAnalysis():
    
    'Set WS to represent Worksheet
   Dim ws As Worksheet
   
    'Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        
        ws.Activate
        
        Dim lastrow As Long
        Dim i, g, a, k As Integer
        Dim Ticker As String
        Dim openingPrice, closingPrice, Quarterchange, TotalVolume As Double
        Dim GreatestIncrease, GreatestDecrease, GreatestVolume As Double
        Dim Maxrow, MinRow, maxvolumeRow As Integer
        
        'Set Last row
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        
        TotalVolume = 0 'Set totalvolume to zero
        k = 2 'set k to two, k represent the output row
        a = 0 'set a to zero, a represent the start of a new ticker
        
        'Headers for output columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("l1").Value = "Total Stock Volume"
        
        'loop through each rows in the dataset starting from the second row
        For i = 2 To lastrow
            'loop for column 1 reference, which holds the Tickers
            For g = 1 To 1
            
                'check if the ticker in the next row is different
                 If ws.Cells(i + 1, g).Value <> ws.Cells(i, g).Value Then
                    
                    Ticker = ws.Cells(i, g).Value ' Ticker
                    TotalVolume = TotalVolume + ws.Cells(i, 7).Value ' Total Volume
                    closingPrice = ws.Cells(i, 6).Value ' Close Price
                    
                    'Calculate Quarter change and Percentage change
                    Quarterchange = closingPrice - openingPrice
                    Percentchange = Quarterchange / openingPrice
                    
                    'Output the Ticker, Total Volume, Quarterchange and Percent change
                    ws.Cells(k, 9).Value = Ticker
                    ws.Cells(k, 12).Value = TotalVolume
                    ws.Cells(k, 10).Value = Quarterchange
                    ws.Cells(k, 11).Value = Percentchange
                    
                        'Conditional formating for Quarterly change and Percent Change
                        If Quarterchange > 0 Then
                             ws.Cells(k, 10).Interior.Color = RGB(127, 255, 0)
                             ws.Cells(k, 11).Interior.Color = RGB(127, 255, 0)
                        ElseIf Quarterchange < 0 Then
                             ws.Cells(k, 10).Interior.Color = RGB(225, 0, 0)
                             ws.Cells(k, 11).Interior.Color = RGB(225, 0, 0)
                            
                        End If
                
                    'format the value of Percent change column as a percent
                    ws.Range("k" & k).NumberFormat = "0.00%"
                    
                    
                    k = k + 1 'add 1 to the row number
                    TotalVolume = 0 'Set the Totalvolume to zero
                    a = 0 'set a to zero
                    
                  'If ticker is the same else statement will activate
                  Else
                    'add the volumes until ticker is changed
                    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                    
                        'Check if a new ticker started to get the opening price
                        If a < 1 Then
                            openingPrice = ws.Cells(i, 3).Value 'Open Price
                            a = 1
                        End If
                        
                    End If
                    
              Next g
              
            Next i
            
            'Headers for calculated Value outputs
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            
            'Used Max function to get the greatest percent increase from Percent change column
            GreatestIncrease = Application.WorksheetFunction.Max(Range("K2:k" & lastrow))
            
            'Identify GreatestIncrease row number and set that to Maxrow
            Maxrow = Application.WorksheetFunction.Match(GreatestIncrease, ws.Columns(11), 0)
           
            'Used Min function to get the greatest percent decrease from Percent change column
            GreatestDecrease = Application.WorksheetFunction.Min(Range("K2:k" & lastrow))
            
            'Identify GreatestDecrease row number and set that to Minrow
            MinRow = Application.WorksheetFunction.Match(GreatestDecrease, ws.Columns(11), 0)
            
            'Used Max function to get the greatest percent increase from Total Stock Volume column
            GreatestVolume = Application.WorksheetFunction.Max(Range("L2:L" & lastrow))
            
            'Identify GreatestVolume row and set tha to MaxvolumeRow
            maxvolumeRow = Application.WorksheetFunction.Match(GreatestVolume, ws.Columns(12), 0)
            
            'Print output values for Greatest % Increase, Greatest % Decrease and Greatest Total Volume
            ws.Range("Q2").Value = GreatestIncrease
                ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").Value = GreatestDecrease
                ws.Range("Q3").NumberFormat = "0.00%"
            ws.Range("Q4").Value = GreatestVolume
            

            ws.Range("P2").Value = ws.Cells(Maxrow, 9).Value ' Greatest % increase Ticker
            ws.Range("P3").Value = ws.Cells(MinRow, 9).Value ' Greatest % Decrease Ticker
            ws.Range("P4").Value = ws.Cells(maxvolumeRow, 9).Value ' Greatest total volume Ticker
            
          Next ws
          
End Sub
