Option Explicit



Sub StockDataAnalysis()

    Dim i, pos_i, pos_j, FirstRecord_index, RecCount As Long
    Dim OpenPrice, ClosePrice, YearlyChange As Double
    
    Dim TickerSymbol As String
    Dim ws As Worksheet
    
    'Loop through all the sheets of the workbook
    For Each ws In ThisWorkbook.Worksheets
    
        ws.Activate
   
        'total records in the worksheet
        RecCount = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
       'Assigning column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
        'format yearly change column as Percentage
        ws.Range("K:K").NumberFormat = "0.00%"
        
        'format greatest % increase and greatest % decrease cells as Percentage
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    
        'Initial assignment of row and column number
        pos_i = 2
        pos_j = 9
        
        'holds the first record row index for a stock
        FirstRecord_index = 2
        
    
        TickerSymbol = ws.Cells(2, 1).Value
       
        
        'Total stock volume
        ws.Cells(pos_i, 12).Value = 0
    
    
        For i = 2 To RecCount
    
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                TickerSymbol = ws.Cells(i, 1).Value
                
                ws.Cells(pos_i, pos_j) = TickerSymbol
                
                'Total Stock Volume
                'adding the cell values instead of using a variable due to large value
                ws.Cells(pos_i, 12).Value = ws.Cells(pos_i, 12).Value + ws.Cells(i, 7).Value
                
                'Calculate yearly change
                
                OpenPrice = ws.Cells(FirstRecord_index, 3).Value
                ClosePrice = ws.Cells(i, 6).Value
                YearlyChange = ClosePrice - OpenPrice
                ws.Cells(pos_i, 10).Value = YearlyChange
             
                'Calculate percentage change
                ws.Cells(pos_i, 11).Value = YearlyChange / OpenPrice
             
             
                'Formatting the cell colors of yearly change column
                If ws.Cells(pos_i, 10).Value >= 0 Then
    
                    ws.Cells(pos_i, 10).Interior.ColorIndex = 4
         
                Else
    
                    ws.Cells(pos_i, 10).Interior.ColorIndex = 3
                End If
                
                
                'prepare for the next stock
                pos_i = pos_i + 1
             
                If i < RecCount Then
                
                    'Reset Total Stock Volume to 0
                     ws.Cells(pos_i, 12).Value = 0
                
                End If
                
                'Row number of the first record of next stock
                FirstRecord_index = i + 1
                
               
            Else
               
                'adding the cell values instead of using a variable due to large value
                 ws.Cells(pos_i, 12).Value = ws.Cells(pos_i, 12).Value + ws.Cells(i, 7).Value
                
                       
            End If
        
         Next i
    
     
       
       
     
        'Greatest % increase
        ws.Cells(2, 17).Value = Application.WorksheetFunction.Max(ws.Range("K:K"))
        
        
        'Greatest % decrease
        ws.Cells(3, 17).Value = Application.WorksheetFunction.Min(ws.Range("K:K"))
        
        'Greatest total volume
        ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Range("L:L"))
    
    
    
        'Findng the ticker symbol for greatest % increase, greatest % decrease and greatest total Volume
        'pos_i holds the last row number of the total Stock volume summary
        
        For i = 2 To pos_i

            If ws.Cells(i, 11) = Application.WorksheetFunction.Max(ws.Range("K:K")) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value

            End If


            If ws.Cells(i, 11) = Application.WorksheetFunction.Min(ws.Range("K:K")) Then
                 ws.Cells(3, 16).Value = ws.Cells(i, 9).Value

            End If


            If ws.Cells(i, 12) = Application.WorksheetFunction.Max(ws.Range("L:L")) Then
                 ws.Cells(4, 16).Value = ws.Cells(i, 9).Value

            End If



        Next i
    
    Next ws
    
End Sub

