Attribute VB_Name = "Module1"
Sub stocks()
    For Each ws In Worksheets
        
        Dim openValue As Double
        Dim closeValue As Double
        
        Dim quarterlyChange As Double
        Dim percentChange As Double
        'declare ticker unique name
        Dim totalstockvolume As Double
        Dim tickerName As String
        'index to keep track of rows in summary table
        Dim tickerIndex As Integer
        
        Dim maxIncrease As Double
        Dim maxDecrease As Double
        Dim maxVolume As Double
        
        maxIncrease = 0
        maxDecrease = 0
        maxVolume = 0
        totalstockvolume = 0
        tickerIndex = 2
        'insert last name variable as reference for for loop
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'column and header creation
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
        'initial ticker and price assignment for first row
        ws.Cells(tickerIndex, 9).Value = Cells(tickerIndex, 1).Value
        openValue = ws.Cells(tickerIndex, 3).Value
        ws.Cells(tickerIndex, 10) = openValue
        'loop through each row to check ticker
        For i = 2 To lastrow
            
            'detect change of ticker
            If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                'obtain ticker name from data
                tickerName = ws.Cells(i, 1).Value
                'adds final row of each ticker set to total volume
                totalstockvolume = totalstockvolume + Cells(i, 7).Value
                closeValue = ws.Cells(i, 6).Value
                'apply ticker name to summary list and  obtain close value
                ws.Cells(tickerIndex, 9).Value = tickerName
                
                ws.Cells(tickerIndex, 12).Value = totalstockvolume
                'reset stock volume over again
                totalstockvolume = 0
                quarterlyChange = closeValue - openValue
                'quarterly and percent changes calculated and assigned to cells, conditional formatting for positive or negative net changes
                ws.Cells(tickerIndex, 10).Value = quarterlyChange
                If (ws.Cells(tickerIndex, 10) > 0) Then
                    ws.Cells(tickerIndex, 10).Interior.ColorIndex = 4
                    
                ElseIf (ws.Cells(tickerIndex, 10) < 0) Then
                    ws.Cells(tickerIndex, 10).Interior.ColorIndex = 3
                    
                End If
                percentChange = quarterlyChange / openValue
                
                ws.Cells(tickerIndex, 11).Value = percentChange
                
                'formatting style change for percent column
                ws.Cells(tickerIndex, 11).NumberFormat = "0.00%"
                
                'find next open stock price for the next descending ticker
                openValue = ws.Cells(i + 1, 3).Value
                
                'descend ticker one row
                tickerIndex = tickerIndex + 1
                
                
            Else
            
            'increase stock volume if ticker name is the same
                totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
                
            End If
            
        Next i
        'looping through new table to find greatest percent increase
        For e = 2 To 1501
            If ws.Cells(e, 11).Value > maxIncrease Then
    
                greatest_ticker = ws.Cells(e, 9).Value
                maxIncrease = ws.Cells(e, 11).Value
       
        End If
    
        Next e
        
        'placing obtained values into worksheet
        ws.Cells(2, 16).Value = maxIncrease
        ws.Cells(2, 15).Value = greatest_ticker
        
        ws.Range("P2").NumberFormat = "0.00%"
        
        'looping to find greatest percent decrease
        
        For e = 2 To 1501
        If ws.Cells(e, 11).Value < maxDecrease Then
            lowest_ticker = ws.Cells(e, 9).Value
            maxDecrease = ws.Cells(e, 11).Value
       
        End If
    
        Next e
        
        ws.Cells(3, 16).Value = maxDecrease
        ws.Cells(3, 15).Value = lowest_ticker
        'formatting percent value
        ws.Range("P3").NumberFormat = "0.00%"
        
        ' looping to find greatest total volume
        
        For e = 2 To 1501
            If ws.Cells(e, 12) > maxVolume Then
                highest_ticker = ws.Cells(e, 9).Value
                maxVolume = ws.Cells(e, 12).Value
            End If
        ws.Cells(4, 15).Value = highest_ticker
        ws.Cells(4, 16).Value = maxVolume
   Next e

Next ws


End Sub
