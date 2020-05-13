Attribute VB_Name = "Module1"
' Production code - Stock Analyzer
Sub WSStockAnalyzer()

	'Visit every worksheet
    For Each ws In Worksheets
        'Get values for the last row and the last column of data
        wsLastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        wsLastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

        'Initialize process variables
        wsWriteRow = 2
        wsWriteColumn = wsLastColumn + 2
        startingColumn = 1
        totalTickerVolume = 0
        firstTickerRow = True

        'Initialize variables for greatest changes
        GreatestPercentIncrease = 0
        GreatestPercentDecrease = 0
        greatestTotalVolume = 0
        greatestPercentIncreaseTicker = ""
        greatestPercentDecreaseTicker = ""
        greatestTotalVolumeTicker = ""

        'Print titles for tickers' summary information
        ws.Cells(1, wsWriteColumn).Value = "Ticker"
        ws.Cells(1, wsWriteColumn + 1).Value = "Yearly Change"
        ws.Cells(1, wsWriteColumn + 2).Value = "Percent Change"
        ws.Cells(1, wsWriteColumn + 3).Value = "Total Stock Volume"

        'Read every row of ticker data 
        For i = 2 To wsLastRow
                'Get initial data from the current ticker being read
                currentTickerName = ws.Cells(i, startingColumn).Value
                currentTickerName = Application.WorksheetFunction.Trim(currentTickerName) 'In case the ticker's name has spaces
                nextTickerName = ws.Cells(i + 1, startingColumn).Value
                nextTickerName = Application.WorksheetFunction.Trim(nextTickerName) 'In case the ticker's name has spaces
                tickerVolume = ws.Cells(i, 7).Value
                totalTickerVolume = totalTickerVolume + tickerVolume

                tickerDate = ws.Cells(i, 2).Value
                
                'Is it the first row of data? Get the ticker's opening price at the beginning of the year
                If firstTickerRow Then
                    firstOpen = ws.Cells(i, 3)
                    firstTickerRow = False
                End If

                'Ticker's last row. Print summary information about the ticker.
                If currentTickerName <> nextTickerName Then

                    'Closing price at the end of the year
                    lastClose = ws.Cells(i, 6)

                    ' Print ticker's name
                    ws.Cells(wsWriteRow, wsWriteColumn).Value = currentTickerName

                    'Print and format yearly change between the opening price and the closing price
                    yearlyChange = lastClose - firstOpen
                    ws.Cells(wsWriteRow, wsWriteColumn + 1).Value = yearlyChange
                    If yearlyChange > 0 Then
                        ws.Cells(wsWriteRow, wsWriteColumn + 1).Interior.ColorIndex = 4
                    Else
                        ws.Cells(wsWriteRow, wsWriteColumn + 1).Interior.ColorIndex = 3
                    End If

                    'Verify that there is no division by zero, to avoid a mathematical error
                    If (lastClose <> 0 And firstOpen <> 0) Then
                        percentChange = (lastClose / firstOpen) - 1
                    Else
                        percentChange = 0
                    End If

                    'Print and format  percent change between opening anc losing price
                    ws.Cells(wsWriteRow, wsWriteColumn + 2).Value = Format(percentChange, "0.00%")
                    If percentChange > 0 Then
                        ws.Cells(wsWriteRow, wsWriteColumn + 2).Interior.ColorIndex = 4
                        'Keep track of the ticker with the greatest percent increase change
                        If percentChange > GreatestPercentIncrease Then
                            GreatestPercentIncrease = percentChange
                            greatestPercentIncreaseTicker = currentTickerName
                        End If
                    Else
                        ws.Cells(wsWriteRow, wsWriteColumn + 2).Interior.ColorIndex = 3
                        'Keep track of the ticker with the greatest percent decrease change
                        If percentChange < GreatestPercentDecrease Then
                            GreatestPercentDecrease = percentChange
                            greatestPercentDecreaseTicker = currentTickerName
                        End If
                    End If

                    'Print the ticker's total volume
                    ws.Cells(wsWriteRow, wsWriteColumn + 3).Value = totalTickerVolume

                    'Keep track of the ticker's with the greastest total volume change
                    If totalTickerVolume > greatestTotalVolume Then
                        greatestTotalVolume = totalTickerVolume
                        greatestTotalVolumeTicker = currentTickerName
                    End If

                    'Initialize values for next ticker's data
                    totalTickerVolume = 0
                    wsWriteRow = wsWriteRow + 1
                    firstTickerRow = True

                End If
                
        Next i

        'Print summary information about the ticker's with the greates changes
        'Print titles
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % increase"
        ws.Cells(3, 14).Value = "Greatest % decrease"
        ws.Cells(4, 14).Value = "Greatest total volume"
        'Print values
        ws.Cells(2, 15).Value = greatestPercentIncreaseTicker
        ws.Cells(3, 15).Value = greatestPercentDecreaseTicker
        ws.Cells(4, 15).Value = greatestTotalVolumeTicker
        ws.Cells(2, 16).Value = Format(GreatestPercentIncrease, "0.00%")
        ws.Cells(3, 16).Value = Format(GreatestPercentDecrease, "0.00%")
        ws.Cells(4, 16).Value = Format(greatestTotalVolume, "General Number")

    Next ws
    
End Sub
