Sub Stocks()

    For Each ws In Worksheets
        ' Initialize headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        Dim Ticker As String
        Dim QuarterlyChange As Double
        Dim PercentChange As Double
        Dim Volume As Double
        Dim StockOpen As Double
        Dim StockClose As Double
        
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Volume = 0

        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2

        ' Loop through each row of stock data
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                Volume = Volume + ws.Cells(i, 7).Value

                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("L" & Summary_Table_Row).Value = Volume

                ' Calculate Quarterly Change and Percent Change
                StockClose = ws.Cells(i, 6).Value
                If StockOpen <> 0 Then
                    QuarterlyChange = StockClose - StockOpen
                    PercentChange = (StockClose - StockOpen) / StockOpen
                Else
                    QuarterlyChange = 0
                    PercentChange = 0
                End If
                
                ws.Range("J" & Summary_Table_Row).Value = QuarterlyChange
                ws.Range("K" & Summary_Table_Row).Value = PercentChange
                ws.Range("K" & Summary_Table_Row).Style = "Percent"
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

                Summary_Table_Row = Summary_Table_Row + 1
                Volume = 0 ' Reset volume for next ticker
            ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                StockOpen = ws.Cells(i, 3).Value
            Else
                Volume = Volume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Variables to track greatest increase, decrease, and volume
        Dim greatestIncrease As Double, greatestDecrease As Double, greatestVolume As Double
        Dim increaseName As String, decreaseName As String, greatestName As String

        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0

        ' Loop through summary data to find greatest changes
        For K = 2 To Summary_Table_Row - 1
            Dim current_k As Double
            Volume = ws.Cells(K, 12).Value
            current_k = ws.Cells(K, 11).Value

            ' Check for greatest increase
            If current_k > greatestIncrease Then
                greatestIncrease = current_k
                increaseName = ws.Cells(K, 9).Value
            End If

            ' Check for greatest decrease
            If current_k < greatestDecrease Then
                greatestDecrease = current_k
                decreaseName = ws.Cells(K, 9).Value
            End If

            ' Check for greatest volume
            If Volume > greatestVolume Then
                greatestVolume = Volume
                greatestName = ws.Cells(K, 9).Value
            End If
        Next K

        ' Output summary results
        ws.Range("N1").Value = "Column Name"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker Name"
        ws.Range("P1").Value = "Value"

        ws.Range("O2").Value = increaseName
        ws.Range("O3").Value = decreaseName
        ws.Range("O4").Value = greatestName
        ws.Range("P2").Value = greatestIncrease
        ws.Range("P3").Value = greatestDecrease
        ws.Range("P4").Value = greatestVolume

        ws.Range("P2:P3").NumberFormat = "0.00%"

        ' Color the changes for visual representation
        For i = 2 To Summary_Table_Row - 1
            If ws.Range("J" & i).Value >= 0 Then
                ws.Range("J" & i).Interior.ColorIndex = 4 ' Green
            Else
                ws.Range("J" & i).Interior.ColorIndex = 3 ' Red
            End If

            If ws.Range("K" & i).Value >= 0 Then
                ws.Range("K" & i).Interior.ColorIndex = 4 ' Green
            Else
                ws.Range("K" & i).Interior.ColorIndex = 3 ' Red
            End If
        Next i
    Next ws

End Sub
