Attribute VB_Name = "Module1"
Sub Stock_Ticker()
    Dim WS As WorkSheet
    For Each WS In Worksheets
        WS.Activate
        
        'Print tickers in new column'
        Dim ticker As String
        Dim nextRow As Integer
        
        Dim volTotal As LongLong
        Dim openPrice As Double
        Dim closePrice As Double
        Dim yearlyChange As Double
        
        Dim greatInc As Double
        Dim greatDec As Double
        Dim greatTotalVol As Double
        
        nextRow = 2
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        greatInc = 0
        greatDec = 0
        greatTotalVol = 0
        
        'Add new column headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        'Loop through each row in first column'
        For i = 2 To lastRow
            'check if previous ticker is different than current
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                'First row with current ticker, get open value
                openPrice = Cells(i, 3).Value
            End If
            
            volTotal = CLngLng(Cells(i, 7).Value) + volTotal
            
            'check if next ticker is different than current
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                    'Assign current ticker to new column
                    ticker = Cells(i, 1).Value
                    'Print variable on new cell'
                    Cells(nextRow, 9).Value = ticker
                    
                    'Last Row w/ current ticker, get close value
                    closePrice = Cells(i, 6).Value
                    'Print Yearly Change
                    yearlyChange = closePrice - openPrice
                    Cells(nextRow, 10).Value = yearlyChange
                    
                    If yearlyChange < 0 Then
                        'Negative change, change cell to RED
                        Cells(nextRow, 10).Interior.ColorIndex = 3
                        Cells(nextRow, 11).Value = Format((closePrice / openPrice - 1), "Percent")
                    ElseIf yearlyChange >= 0 Then
                        'Positive/no change, change cell to GREEN
                        Cells(nextRow, 10).Interior.ColorIndex = 4
                        Cells(nextRow, 11).Value = Format((closePrice / openPrice - 1), "Percent")
                    End If
                    
                    'Add totalVol to new Cell column
                    Cells(nextRow, 12).Value = volTotal
                    'Reset volTotal for next ticker
                    volTotal = 0
                    
                    'increment counter'
                    nextRow = nextRow + 1
            End If
        Next i
        
        For j = 2 To lastRow
            currentPercent = Cells(j, 11).Value
            
            'Find Greatest Increase
            If currentPercent > greatInc Then
                greatInc = currentPercent
                Cells(2, 16).Value = Cells(j, 9).Value
                Cells(2, 17).Value = Format(Cells(j, 11).Value, "Percent")
            End If
            
            'Find Greatest Decrease
            If currentPercent < greatDec Then
                greatDec = currentPercent
                Cells(3, 16).Value = Cells(j, 9).Value
                Cells(3, 17).Value = Format(Cells(j, 11).Value, "Percent")
            End If
            
            currentNum = Cells(j, 12).Value
            
            'Find Greatest Total
            If currentNum > greatTotalVol Then
                greatTotalVol = currentNum
                Cells(4, 16).Value = Cells(j, 9).Value
                Cells(4, 17).Value = Cells(j, 12).Value
            End If
        Next j
    Next WS
End Sub



