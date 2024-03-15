# Module_2_Challenge
Module 2 Challenge 

## This is the README for the Module 2 Challenge for Sultan Raheem

NOTE: This assignment has been done with group work with myself, Rob Molenda, Lucas Perez, and Muneeb Samad
I also got verbal input from Cadeem Musgrove separately on Zoom regarding conditional formatting.

The whole of our projects have been achieved through teamwork as a cohort.

## Within the submission there are screenshots for 2018, 2019, and 2020 results in the stock data

## There is also a BAS file for the VBA script and it is called Sub Stocks_Test

## This submission creates a script that loops and outputs results for all the stocks, it takes the Ticker name, Yearly Change, Percent Change, and Total Stock Volume

This is achieved by looking at the solutions for grading within the class and the lotto numbers solution done in class

the relevant lines of code are: 

' Loop through each worksheet
    For Each ws In Worksheets

        Next ws ' Close the loop through worksheets

## It also calculates the Greatest Percent Increase, Decrease, and Greatest Total Volume

These were helped achieved by Rob during our group work, the lines of code are:

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

## It does this for all three sheets for 2018 to 2020

This is because we use ws as a reference and we have already specified

' Loop through each worksheet
    For Each ws In Worksheets

        Next ws ' Close the loop through worksheets

        Further Source: https://stackoverflow.com/questions/52012092/vba-loop-of-multiple-sheets-in-a-worksheet

        and a source from Rob: https://www.mrexcel.com/board/threads/repeating-excel-macro-across-multiple-worksheets.1028802/

## There are also calculations and variables set for the ticker name, volume of stock, open/close price

Relevant code:

Dim TickerName As String
        Dim StockVolume As Double
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim TotalVol As Double

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

                Printing was achieved like this:

                                ws.Range("K" & Summary_Table_Row).Value = Percentchange
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                ws.Range("L" & Summary_Table_Row).Value = TotalVol

                                ws.Range("I" & Summary_Table_Row).Value = TickerName
                ws.Range("J" & Summary_Table_Row).Value = YearlyChange

                And as far as all percent changes formats, the source used was: https://www.statology.org/vba-percentage-format/

## columns are also created for tickers, total volume, yearly change, and percent change through the script, further are created for greatest increase, decrease and greatest total volume through the VBA script
These were done with these lines of code:

        ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
        ws.Range("P1:Q1").Value = Array("Ticker", "Value")
        ws.Range("O2").Value = Array("Greatest % Increase")
        ws.Range("O3").Value = Array("Greatest % Decrease")
        ws.Range("O4").Value = Array("Greatest Total Volume")

  These concepts were learned from the source: https://www.mrexcel.com/board/threads/add-column-headers-in-a-worksheet-using-vba.1078803/

## Conditional formatting was also done in VBA to highlight the colours in the yearly change column

Concepts for conditional formatting were verbally understood by Cadeem Musgrove on a Zoom call, and later implemented with my cohort as a team

                A source to help grasp this concept include: 
                https://learn.microsoft.com/en-us/office/vba/api/excel.colorindex
                https://www.excel-pratique.com/en/vba_tricks/cell-color-conditional-formatting

This was achieved with a combination of help from source and cohort, see below comment and sources for more conditional formatting.

## This conditional formatting and all calculations ran for all three of the sheets that are in the Excel file

this is with the aforementioned worksheet code and using ws. within the VBA Script

Relevant lines of code:

                If YearlyChange > 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If

                A source to help grasp this concept include: 
                https://learn.microsoft.com/en-us/office/vba/api/excel.colorindex
                https://www.excel-pratique.com/en/vba_tricks/cell-color-conditional-formatting

## One thing missing from this submission is the name of the ticker associated with the greatest % increase/decrease and greatest total volume

I was not able to understand how to achieve this part of the solution

-----------------------------------------------------------------------------------------------------

All mentioned sources and further sources are in a separate document & below:

https://www.mrexcel.com/board/threads/add-column-headers-in-a-worksheet-using-vba.1078803/

https://stackoverflow.com/questions/57367032/how-can-i-select-a-cell-given-its-row-and-column-number

https://www.statology.org/vba-percentage-format/

https://stackoverflow.com/questions/52012092/vba-loop-of-multiple-sheets-in-a-worksheet

https://www.excel-pratique.com/en/vba_tricks/cell-color-conditional-formatting

https://learn.microsoft.com/en-us/office/vba/api/excel.colorindex

https://www.mrexcel.com/board/threads/repeating-excel-macro-across-multiple-worksheets.1028802/

https://www.reddit.com/r/vba/comments/rek9i0/what_does_i_j_and_k_mean/

https://stackoverflow.com/questions/4137785/why-are-variables-i-and-j-used-for-counters/454308#454308

https://trumpexcel.com/vba-loops/

https://trumpexcel.com/vba-worksheets/

https://www.indeed.com/career-advice/career-development/how-to-enable-macros-in-excel#:~:text=To%20find%20%22Macro%20Settings%2C%22,OK%22%20to%20enable%20all%20macros.

https://www.reddit.com/r/vba/comments/fq5e0r/i_keep_getting_a_compile_error_next_without_for/

https://chat.openai.com/share/1dc70cdb-2689-4aaa-a3c2-874a361d4cf9
