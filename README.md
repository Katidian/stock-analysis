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

The purpose of the TickerIndex variable is to tell us which ticker index number we're working with at any given time. I thought that instead of starting with a For 
loop that loops through each ticker in turn, we need to start with a For loop that loops through all the rows of data.

But then I got stuck on instruction 2a ("Create a for loop to initialize the TickerVolumes to zero") because I wasn't quite sure what I should be looping through. After thinking about this for way too long, I decided that the point is to reset the TickerVolumes every time we move to a new TickerIndex number. Doing this 
does not actually run through the data rows 12 times, since it's separate from the For loop that WILL run the data rows.

```
    For TickerIndex = 0 To 11
    
        TickerVolumes(TickerIndex) = 0
    
    Next i
```

The instructions for the challenge did not specify where to set the Ticker variable, but I assume we still need it, so I set it right after creating the 
TickerIndex but before creating the three output arrays.

```
    '1a) Create a ticker Index
    'and set it to zero
    TickerIndex = 0
    
    'I think we need to set the Ticker variable, too.
    Ticker = Tickers(TickerIndex)

    '1b) Create three output arrays
    Dim TickerVolumes(TickerIndex) As Long
    Dim TickerStartingPrices(TickerIndex) As Single
    Dim TickerEndingPrices(TickerIndex) As Single
```

But VBA did not like this. When I tested my code by stepping through it using the debugging tool, I got an error about this line:

```
'1b) Create three output arrays
    Dim TickerVolumes(TickerIndex) As Long
    Dim TickerStartingPrices(TickerIndex) As Single
    Dim TickerEndingPrices(TickerIndex) As Single
```


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

