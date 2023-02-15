Sub summarizeStockData()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    Call populate_sheet(ws)
Next ws

End Sub

Sub populate_sheet(ws As Worksheet)

    Dim table As Range
    Dim i As Integer
    Dim startPrice As Double
    Dim endPrice As Double
    Dim priceChange As Double
    Dim totalVol As Double
    Dim maxPriceInc As Double
    Dim maxPriceDec As Double
    Dim maxVol As Double
    Dim maxPriceIncTicker As String
    Dim maxPriceDecTicker As String
    Dim maxVolTicker As String

    ws.Range("J1").Value = "Ticker"
    ws.Range("J1").EntireColumn.AutoFit
    ws.Range("K1").Value = "Yearly Change"
    ws.Range("K1").EntireColumn.AutoFit
    ws.Range("L1").Value = "Percent Change"
    ws.Range("L1").EntireColumn.AutoFit
    ws.Range("M1").Value = "Total Stock Volume"
    ws.Range("M1").EntireColumn.AutoFit
    ws.Range("Q1").Value = "Ticker"
    ws.Range("Q1").EntireColumn.AutoFit
    ws.Range("R1").Value = "Value"
    ws.Range("R1").EntireColumn.AutoFit
    ws.Range("P2").Value = "Greatest % Increase"
    ws.Range("P2").EntireColumn.AutoFit
    ws.Range("P3").Value = "Greatest % Decrease"
    ws.Range("P3").EntireColumn.AutoFit
    ws.Range("P4").Value = "Greatest Total Volume"
    ws.Range("P4").EntireColumn.AutoFit

    i = 2
    Set table = ws.Range("A2", ws.Range("A2").End(xlToRight).End(xlDown))
   
    
    For Each Row In table.Rows
        totalVol = totalVol + CDbl(Row.Cells(1, 7).Value)
        If Row.Cells(1, 2).Value = ws.Name & "0102" Then
            startPrice = CDbl(Row.Cells(1, 3).Value)
        ElseIf Row.Cells(1, 2).Value = ws.Name & "1231" Then
            endPrice = CDbl(Row.Cells(1, 6).Value)
            priceChange = endPrice - startPrice
            ws.Range("J" & i).Value = Row.Cells(1, 1).Value
            ws.Range("K" & i).Value = priceChange
            ws.Range("L" & i).Value = FormatPercent(priceChange / startPrice)
            ws.Range("M" & i).Value = totalVol
            If (priceChange < 0) Then
                ws.Range("K" & i).Interior.Color = RGB(255, 0, 0)
                If (priceChange < maxPriceDec) Then
                    maxPriceDec = priceChange
                    maxPriceDecTicker = Row.Cells(1, 1).Value
                End If
            Else
                ws.Range("K" & i).Interior.Color = RGB(0, 255, 0)
                If (priceChange > maxPriceInc) Then
                    maxPriceInc = priceChange
                    maxPriceIncTicker = Row.Cells(1, 1).Value
                End If
            End If
            If (totalVol > maxVol) Then
                maxVol = totalVol
                maxVolTicker = Row.Cells(1, 1).Value
            End If
            totalVol = 0
            i = i + 1
        End If
    
    
    Next
    
    ws.Range("Q2").Value = maxPriceIncTicker
    ws.Range("R2").Value = maxPriceInc
    ws.Range("Q3").Value = maxPriceDecTicker
    ws.Range("R3").Value = maxPriceDec
    ws.Range("Q4").Value = maxVolTicker
    ws.Range("R4").Value = maxVol

End Sub



