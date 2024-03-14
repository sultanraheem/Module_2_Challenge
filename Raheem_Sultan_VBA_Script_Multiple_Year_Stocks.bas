Attribute VB_Name = "Module1"
Sub Stocks_Test()

    ' Loop through each worksheet
    For Each ws In Worksheets

        ' Set column names with array
        ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")

' titles made for new criteria
    ws.Range("P1:Q1").Value = Array("Ticker", "Value")
        ws.Range("O2").Value = Array("Greatest % Increase")
        ws.Range("O3").Value = Array("Greatest % Decrease")
        ws.Range("O4").Value = Array("Greatest Total Volume")

        Dim TickerName As String
        Dim StockVolume As Double
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim TotalVol As Double
        Dim YearlyChange As Double
        Dim Percentchange As Double
        Dim Summary_Table_Row As Double
        Dim GreatestPercentIncrease As Double
        Dim GreatestPercentDecrease As Double
        Dim GreatestTotalVolume As Double
        
        

        ' Initialize variables
        StockVolume = 0
        OpenPrice = 0
        ClosePrice = 0
        TotalVol = 0
        Summary_Table_Row = 2 ' Avoid header row

        ' Last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Loop through rows
        For i = 2 To LastRow ' Skip header row

            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                OpenPrice = ws.Cells(i, 3).Value ' Set OpenPrice
            End If

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                TickerName = ws.Cells(i, 1).Value
                ClosePrice = ws.Cells(i, 6).Value
                StockVolume = StockVolume + ws.Cells(i, 3).Value

                YearlyChange = ClosePrice - OpenPrice

                If OpenPrice = 0 Then
                    Percentchange = 0
                Else
                    Percentchange = (ClosePrice - OpenPrice) / OpenPrice
                End If

                TotalVol = TotalVol + ws.Cells(i, 7).Value

                ' Print results to summary row table
                ws.Range("I" & Summary_Table_Row).Value = TickerName
                ws.Range("J" & Summary_Table_Row).Value = YearlyChange

                If YearlyChange > 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If

                ws.Range("K" & Summary_Table_Row).Value = Percentchange
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                ws.Range("L" & Summary_Table_Row).Value = TotalVol

                Summary_Table_Row = Summary_Table_Row + 1

                OpenPrice = 0
                ClosePrice = 0
                TotalVol = 0
            Else
                TotalVol = TotalVol + ws.Cells(i, 7).Value
            End If

        Next i ' Close the loop through rows
        
        Dim rng As Range
    Dim maxvalue As Variant
    Set rng = ws.Range("K:K")
    maxvalue = Application.WorksheetFunction.Max(rng)


    ws.Range("Q2").Value = maxvalue
    ws.Range("Q2").NumberFormat = "0.00%"

    Dim minvalue As Variant
    Set rng = ws.Range("K:K")
    minvalue = Application.WorksheetFunction.Min(rng)

    ws.Range("Q3").Value = minvalue
    ws.Range("Q3").NumberFormat = "0.00%"

    Set rng = ws.Range("L:L")
    maxvalue = Application.WorksheetFunction.Max(rng)


    ws.Range("Q4").Value = maxvalue



    Next ws ' Close the loop through worksheets

End Sub

