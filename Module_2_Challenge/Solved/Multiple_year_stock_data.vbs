Sub CalculateStockData()
    Dim WorksheetName As String
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Long
    Dim greatIncrease As Double
    Dim greatDecrease As Double
    Dim greatVolume As Double
    Dim ticker_greatIncrease As String
    Dim ticker_greatDecrease As String
    Dim ticker_greatVolume As String
        
    For Each ws In Worksheets
        WorksheetName = ws.Name
        
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        summaryRow = 2
        
        ' Set initial values for opening price, total volume, Great Increase, and Great Decrease
        openingPrice = ws.Cells(2, 3).Value
        totalVolume = 0
        greatIncrease = 0
        greatDecrease = 0
        
        ' Loop through each row of data
        For i = 2 To lastRow
            ' Check if ticker symbol changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Get closing price and calculate yearly change and percent change
                closingPrice = ws.Cells(i, 6).Value
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                yearlyChange = closingPrice - openingPrice
                    If openingPrice <> 0 Then
                        percentChange = (yearlyChange / openingPrice)
                        ' This is to decide who had the great precent increase
                        If percentChange > greatIncrease Then
                            greatIncrease = percentChange
                            ticker_greatIncrease = ws.Cells(i, 1).Value
                        End If
                        ' This is to decide who had the least precent increase
                        If percentChange < greatDecrease Then
                            greatDecrease = percentChange
                            ticker_greatDecrease = ws.Cells(i, 1).Value
                        End If
                    Else
                        percentChange = 0
                    End If
                ' Output results to summary table
                ws.Cells(summaryRow, 9).Value = ws.Cells(i, 1).Value ' Ticker
                
                If yearlyChange >= 0 Then
                ' Yearly change
                    ws.Cells(summaryRow, 10).Value = yearlyChange
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 4 ' green
                Else
                    ' Yearly change
                    ws.Cells(summaryRow, 10).Value = yearlyChange
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 3
                End If
                ' Percentage change
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
                'Total volume
                ws.Cells(summaryRow, 12).Value = totalVolume
                
                ' Move to next row in summary table
                summaryRow = summaryRow + 1
                
                ' Reset opening price for the next ticker
                openingPrice = ws.Cells(i + 1, 3).Value
                
                ' Reset total volume for the next ticker
                totalVolume = 0
            Else
                ' Accumulate total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
            
            
            ' This is to decide who had the great volume
            If totalVolume > greatVolume Then
                greatVolume = totalVolume
                ticker_greatVolume = ws.Cells(i, 1).Value
            End If
        Next i
        
        ws.Cells(2, 16) = greatIncrease
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(2, 14) = "Greatest % Increase"
        ws.Cells(2, 15) = ticker_greatIncrease
        
        
        ws.Cells(3, 16) = greatDecrease
        ws.Cells(3, 16).NumberFormat = "0.00%"
        ws.Cells(3, 15) = ticker_greatDecrease
        ws.Cells(3, 14) = "Greatest % Decrease"
        
        
        ws.Cells(4, 16) = greatVolume
        ws.Cells(4, 15) = ticker_greatVolume
        ws.Cells(4, 14) = "Greatest Total Volume"
        
        ws.Cells(1, "I") = "Ticker"
        ws.Cells(1, "J") = "Yearly Change"
        ws.Cells(1, "K") = "Percent Change"
        ws.Cells(1, "L") = "Total Stock Volume"
        
        ws.Cells(1, "O") = "Ticker"
        ws.Cells(1, "P") = "Value"
        
        ws.Columns.AutoFit
    Next ws
End Sub


