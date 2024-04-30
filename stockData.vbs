Attribute VB_Name = "Module1"
Sub stockData():

    Dim totalVolume As Double ' variable for holding "Total Stock Volume" (column L) aka running total
    Dim row As Long ' variable for going through rows
    Dim rowCount As Double ' variable for numbers of rows
    Dim yearChange As Double ' variable for holding "Yearly Change" (column J)
    Dim percentChange As Double ' variable for holding "Percent Change" (column K)
    Dim summaryRow As Long ' variable to keep track of location for each Ticker abbv.
    Dim ticker As Long ' variable for start of stock row
    
    For Each ws In Worksheets
    
        ' Column headers in all worksheets
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ' initailizing values
        totalVolume = 0
        yearChange = 0
        summaryRow = 0
        ticker = 2
        
        ' determine last row
        rowCount = ws.Cells(Rows.Count, "A").End(xlUp).row
    
        For row = 2 To rowCount ' Loop through all ticker abbv.
        
            ' check if its still the same ticker abbv.
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
            
                totalVolume = totalVolume + ws.Cells(row, 7).Value
                    
                If totalVolume = 0 Then
                    
                    ' print results in summary table
                    ws.Range("I" & 2 + summaryRow).Value = Cells(row, 1).Value
                    ws.Range("J" & 2 + summaryRow).Value = 0
                    ws.Range("K" & 2 + summaryRow).Value = 0 & "%"
                    ws.Range("L" & 2 + summaryRow).Value = 0
                    
                Else
                    ' find first non zero open value
                    If ws.Cells(ticker, 3).Value = 0 Then
                        For findValue = ticker To row
                        
                            ' check next value does not equal 0
                            If ws.Cells(findValue, 3).Value <> 0 Then
                                ticker = findValue
                                Exit For
                            End If
                        Next findValue
                    End If
                    
                    ' calculate year change
                    yearChange = (ws.Cells(row, 6).Value - ws.Cells(ticker, 3).Value)
                    ' percent change
                    percentChange = yearChange / ws.Cells(ticker, 3).Value
                    
                    ws.Range("I" & 2 + summaryRow).Value = ws.Cells(row, 1).Value
                    ws.Range("J" & 2 + summaryRow).Value = yearChange
                    ws.Range("J" & 2 + summaryRow).NumberFormat = "0.00"
                    ws.Range("K" & 2 + summaryRow).Value = percentChange
                    ws.Range("K" & 2 + summaryRow).NumberFormat = "0.00%"
                    ws.Range("L" & 2 + summaryRow).Value = totalVolume
                    
                    If yearChange > 0 Then
                        ws.Range("J" & 2 + summaryRow).Interior.ColorIndex = 4
                    ElseIf yearChange < 0 Then
                        ws.Range("J" & 2 + summaryRow).Interior.ColorIndex = 3
                    Else
                        ws.Range("J" & 2 + summaryRow).Interior.ColorIndex = 0
                    End If
                    
                End If
                
                ' reset total stock volume / yearly change
                totalVolume = 0
                yearChange = 0
                ' move to next row in summary table
                summaryRow = summaryRow + 1
                
            Else
                ' add to the total stock volume
                totalVolume = totalVolume + ws.Cells(row, 7).Value
            
            End If
        
        Next row
        
        'fill greatest / least summary table
        ws.Range("Q2").Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q3").Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q4").Value = "%" & WorksheetFunction.Max(ws.Range("L2:L" & rowCount)) * 100
        ws.Range("Q4").NumberFormat = "#,###"
        
        increaseNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        ws.Range("P2").Value = ws.Cells(increaseNumber + 1, 9)
        
        decreaseNumber = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        ws.Range("P3").Value = ws.Cells(decreaseNumber + 1, 9)
        
        greatNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount))
        ws.Range("P4").Value = ws.Cells(greatNumber + 1, 9)
        
        ws.Columns("A:Q").AutoFit
        
    Next ws

End Sub
