Attribute VB_Name = "Module1"
Sub StockData()

    For Each ws In Worksheets

        'Variables for all ws

        Dim tickerName As String
        Dim yearlyChange As Double
        Dim percentChange As Double
        Dim totalVolume As Variant
        Dim sumTableRow As Variant
        Dim firstOpen As Variant
        Dim finalClose As Variant
        
        'Part 2 Variables
                Dim greatInc As Variant
                Dim greatDec As Variant
                Dim greatTotVol As Variant
        
       
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Initial values
        sumTableRow = 2
        totalVolume = 0
        firstOpen = ws.Cells(2, 3).Value
        
        ' Pull just the first reference of a ticker name into a variable and track volume
        For i = 2 To lastRow
        
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Collects ticker names once for table
                tickerName = ws.Cells(i, 1).Value
                
                'Values to calculate Yearly and Percent Changes
                finalClose = ws.Cells(i, 6).Value
                
                'Calculations
                yearlyChange = finalClose - firstOpen
                percentChange = (finalClose / firstOpen) - 1
                totalVolume = totalVolume + Cells(i, 7).Value
                
                'create columns with ticker and volumeheaders
                ws.Range("I1").Value = "Ticker"
                ws.Range("J1").Value = "Yearly Change"
                ws.Range("K1").Value = "Percent Change"
                ws.Range("L1").Value = "Total Stock Volume"
                
                
                'populate columns with necessary data
                ws.Range("I" & sumTableRow).Value = tickerName
                ws.Range("J" & sumTableRow).Value = yearlyChange
                ws.Range("K" & sumTableRow).Value = percentChange
                ws.Range("L" & sumTableRow).Value = totalVolume
            
                
                'build summary table
                ws.Range("K" & sumTableRow).NumberFormat = "0.00%"
                sumTableRow = sumTableRow + 1
               
            
                'reset variables for calculation
                totalVolume = 0
                firstOpen = ws.Cells(i + 1, 3).Value
                
                
            Else
                'I think this works by calling the sum from just before it's reset?
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            
                
            End If
            
            
        Next i
        
         'Yearly Change Conditional Formatting Loop
        For j = 2 To sumTableRow

            If ws.Range("J" & j).Value < 0 Then
                ws.Range("J" & j).Interior.ColorIndex = 3
                    
            Else
                ws.Range("J" & j).Interior.ColorIndex = 4
                    
            End If
            
            'Percent Change Conditional Formatting Loop
            
            If ws.Range("K" & j).Value < 0 Then
                ws.Range("K" & j).Interior.ColorIndex = 3
                    
            Else
                ws.Range("K" & j).Interior.ColorIndex = 4
                    
            End If
            
        Next j
    
        'lastSumRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        'Part 2 value calcs
        greatInc = WorksheetFunction.Max(ws.Range("k:k"))
        greatDec = WorksheetFunction.Min(ws.Range("k:k"))
        greatTotVol = WorksheetFunction.Max(ws.Range("l:l"))
                
                
        'Part 2 Tickers - can't get to work
        'greatIncTick = ws.Cells(WorksheetFunction.Offset.greatInc(0, -2).Value).Value
        'greatDecTick = ws.Cells(WorksheetFunction.Offset.greatInc(0, -2).Value).Value
        
        'greatIncTick = greatInc.Offset(0, -2).Value
        'greatDecTick = greatDec.Offset(0, -2).Value
        'greatTotVolTick = greatTotVol.Offset(0, -2).Value
        
                
        'Part 2 Labels
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
                
        'Part 2 Results
        ws.Range("P2").Value = greatIncTick
        ws.Range("P3").Value = greatDecTick
        ws.Range("P4").Value = greatTotVolTick
        ws.Range("Q2").Value = greatInc
        ws.Range("Q3").Value = greatDec
        ws.Range("Q4").Value = greatTotVol
                
        'Part 2 Formatting
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        
    
    
    Next ws


End Sub
