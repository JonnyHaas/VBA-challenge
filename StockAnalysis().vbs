Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, summaryRow As Long
    Dim ticker As String
    Dim totalVolume As Double
    Dim startPrice As Double, endPrice As Double
    Dim yearlyChange As Double, percentChange As Double
    Dim maxIncrease As Double, maxDecrease As Double, maxVolume As Double
    Dim maxIncreaseTicker As String, maxDecreaseTicker As String, maxVolumeTicker As String

    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        summaryRow = 2
        startPrice = 0

        ' Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                endPrice = ws.Cells(i, 6).Value
                totalVolume = totalVolume + ws.Cells(i, 7).Value

                yearlyChange = endPrice - startPrice
                If startPrice <> 0 Then
                    percentChange = yearlyChange / startPrice
                Else
                    percentChange = 0
                End If

                ' Update max values
                If percentChange > maxIncrease Then
                    maxIncrease = percentChange
                    maxIncreaseTicker = ticker
                ElseIf percentChange < maxDecrease Then
                    maxDecrease = percentChange
                    maxDecreaseTicker = ticker
                End If
                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    maxVolumeTicker = ticker
                End If

                ' Output
                With ws
                    .Cells(summaryRow, 9).Value = ticker
                    .Cells(summaryRow, 10).Value = yearlyChange
                    .Cells(summaryRow, 11).Value = percentChange
                    .Cells(summaryRow, 11).NumberFormat = "0.00%"
                    .Cells(summaryRow, 12).Value = totalVolume

                    If percentChange < 0 Then
                        .Cells(summaryRow, 11).Interior.Color = vbRed
                    ElseIf percentChange > 0 Then
                        .Cells(summaryRow, 11).Interior.Color = vbGreen
                    End If
                End With

                summaryRow = summaryRow + 1
                totalVolume = 0
                If i + 1 <= lastRow Then
                    startPrice = ws.Cells(i + 1, 3).Value
                End If
            Else
                If startPrice = 0 Then startPrice = ws.Cells(i, 3).Value
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i

        ' Output greatest values
        ws.Cells(2, 16).Value = maxIncreaseTicker
        ws.Cells(2, 17).Value = maxIncrease
        ws.Cells(2, 17).NumberFormat = "0.00%"

        ws.Cells(3, 16).Value = maxDecreaseTicker
        ws.Cells(3, 17).Value = maxDecrease
        ws.Cells(3, 17).NumberFormat = "0.00%"

        ws.Cells(4, 16).Value = maxVolumeTicker
        ws.Cells(4, 17).Value = maxVolume
    Next ws

    Application.ScreenUpdating = True
End Sub

