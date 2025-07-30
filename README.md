# vba-stock-data-analysis
VBA Stock Data Analysis – Module 2 Challenge

## Project Overview

This project is a VBA-based stock data analyzer developed for the Module 2 Challenge. It automates the analysis of stock data from the years 2018, 2019, and 2020. The macro generates summary statistics and applies conditional formatting for clear visualization.

The script calculates:

- Yearly Change  
- Percent Change  
- Total Stock Volume  
- Greatest Percent Increase  
- Greatest Percent Decrease  
- Greatest Total Volume

## Contributors

This project was completed in collaboration with:

- Sultan Raheem  
- Rob Molenda  
- Lucas Perez  
- Muneeb Samad  

Verbal guidance on conditional formatting was also provided by Cadeem Musgrove during a Zoom session.

## Files Included

- `Sub Stocks_Test.bas` – The main VBA macro file  
- Output screenshots for 2018, 2019, and 2020  

## Features

### Worksheet Looping

The macro dynamically loops through all worksheets using the following structure:

```vba
For Each ws In Worksheets
    ' Logic runs here
Next ws
This enables analysis across all yearly sheets without manual updates.

Summary Calculations
For each stock ticker, the script calculates:

Yearly Change

Percent Change (handling division by zero)

Total Stock Volume

Key variables:

vba
Copy
Edit
Dim TickerName As String
Dim StockVolume As Double
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim TotalVol As Double
Maximum and Minimum Metrics
The macro identifies the greatest percent increase, decrease, and highest volume:

vba
Copy
Edit
maxvalue = Application.WorksheetFunction.Max(ws.Range("K:K"))
minvalue = Application.WorksheetFunction.Min(ws.Range("K:K"))
volumeMax = Application.WorksheetFunction.Max(ws.Range("L:L"))

ws.Range("Q2").Value = maxvalue
ws.Range("Q3").Value = minvalue
ws.Range("Q4").Value = volumeMax
Table Headers and Output
Headers and metric labels are added dynamically:

vba
Copy
Edit
ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
ws.Range("P1:Q1").Value = Array("Ticker", "Value")
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
Conditional Formatting
Color formatting is applied to the Yearly Change column:

Green for positive change

Red for negative change

Implementation:

vba
Copy
Edit
If YearlyChange > 0 Then
    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
Else
    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
End If
Multi-Year Analysis
All logic runs across each of the three sheets representing 2018, 2019, and 2020. The use of ws allows the macro to generalize processing for each worksheet.

Known Limitation
The current version does not identify the ticker name associated with the greatest percent increase, greatest percent decrease, or greatest total volume. This functionality remains to be implemented.

References
Sources consulted during development include:

https://stackoverflow.com/questions/52012092/vba-loop-of-multiple-sheets-in-a-worksheet

https://www.mrexcel.com/board/threads/repeating-excel-macro-across-multiple-worksheets.1028802

https://www.statology.org/vba-percentage-format

https://learn.microsoft.com/en-us/office/vba/api/excel.colorindex

https://www.excel-pratique.com/en/vba_tricks/cell-color-conditional-formatting

https://trumpexcel.com/vba-loops

https://trumpexcel.com/vba-worksheets

How to Run
Open the Excel workbook containing the stock data for 2018 to 2020

Open the VBA editor (Alt + F11)

Import Sub Stocks_Test.bas

Run the macro to process the data and populate results across all sheets
