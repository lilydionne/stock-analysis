# Using Excel VBA To Streamline Stock Analysis
## Overview
Steve is happy with our progress, but what happens if he uses the original code to run against the entire stock market? The code will likely take a long time to process based on the way it's structured. The objective of this project analysis was to automate the code used to review stocks in order to speed up execution times. 

## Results
##### Stock Performance Analysis YoY
When comparing performance year over year, we can clearly see that 2018 was not a good year for environmental stocks as the majority of them saw negative return, with only 2 stocks, "RUN" and "ENPH" seeing positive return, and Steve's parents original stock, "DQ", had the worst return. However, in the previous year "DQ" had the best return in 2017.

![2018 VBA Challenge Return](https://github.com/lilydionne/stock-analysis/blob/main/VBA_Challenge_Return.png)

##### Analysis of Code Execution
The refactored code is more efficient than the original. The original was able to complete 2018 data in 0.88 seconds, compared to the refactored code running in 0.11 seconds for the same year. This highlights the benefits of the refactored code and it's likeihood to process more data in less time than using code with nested for statements.

![2017 VBA Challenge Runtimes](https://github.com/lilydionne/stock-analysis/blob/main/VBA_Challenge_2017.PNG) ![2018 VBA Challenge Runtimes](https://github.com/lilydionne/stock-analysis/blob/main/VBA_Challenge_2018.PNG)

The new refactored code was able to eliminated nested for statements by using the tickerIndex in order to cycle through the tickers array, as well as log the values of volume, starting prices, and ending prices in additional arrays.  
````
tickerIndex = 0

'1b) Create three output arrays
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
    
'2a) Create a for loop to initialize the tickerVolumes to zero.
For i = 0 To 11
tickerVolumes(i) = 0
tickerStartingPrices(i) = 0
tickerEndingPrices(i) = 0
        
Next i
````
Now when we needed to reference the array, we did not use a loop to cycle through but instead used the tickerIndex:
````
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
              
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            '3d) Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
                        
        End If
````

Where at the end of the code, the tickerIndexed increased within the loop same loop allowing the arrays for volume, starting price and ending price to log a value, and at the end of the if statement, that value was inserted into the "AllStocksAnalysis" sheet all at once rather than ticker by ticker.

````
    For i = 0 To 11
        
        Worksheets("AllStocksAnalysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
        
    Next i
````


## Summary
Faster execution times allow us to run the code across additional stocks with the potential to run across the entire stock market for insights into volume and return over the years. The original script utilized nested for statements and, while this worked for our current data set, eliminating the nested for statements would allows fewer loops and quicker analysis, which is why we refactored the code. Now Steve can expand his stock market analysis to help further help his parents and future clients determine which stocks will be the right investment for them.
