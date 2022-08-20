# stock-analysis
Analysis for boot camp VBA module

The original code from the module has us looping through all the rows of stock price data 12 times — once for each ticker — each time finding the relevant metrics for 
that ticker. 

For example, here are the beginning of the For loop and the conditional statement that let us run through each row looking for — and adding up — the total daily volume 
numbers for each of the 12 tickers in our Tickers array.

```     
For i = 0 To 11
    
  Ticker = tickers(i)
  TotalVolume = 0

  Worksheets(yearValue).Activate
    For j = 2 To RowCount
    
      If Cells(j, 1).Value = Ticker Then
               
        TotalVolume = TotalVolume + Cells(j, 8).Value
            
      End If
```      
            
My goal in refactoring the code is to avoid a dozen loops through all the rows of stock price data, which I hope will cut down on the program run time. In order to 
achieve this, we have to find and store the relevant metrics (total volume, starting price and ending price) for each ticker along the way while only looping 
through once. 

It also seems unnecesary to activate the output sheet ("All stocks analysis") twice. The original module code has us activating the output worksheet near the beginning
of the subroutine in order to set up data headers. 

```
Sub AllStocksAnalysis()

    Dim startTime As Single
    Dim endTime As Single

    yearValue = InputBox("For which year would you like to analyze stock performance?")

    startTime = Timer

    'Format the output sheet on the "All stocks analysis" worksheet.
    Worksheets("All stocks analysis").Activate
        
    Range("A1").Value = "All stocks (" + yearValue + ")"
            
    'Set up the data headers
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total daily volume"
    Cells(3, 3).Value = "Return"
```

Then we activate the sheet containing the stock price data ("2017" or "2018", depending on the user's choice of year for analysis) and run through the various
For loops and conditional statements that gather and store the appropriate data. As part of the outer For loop that runs through each ticker, we re-activate the 
output sheet and populate it with the data gathered from the previous sheet for each ticker. 

I would like to see if adding the column headers to the output sheet while we already have it activated during the data population process makes the code less 
bulky and/or speeds up the run time.

```

