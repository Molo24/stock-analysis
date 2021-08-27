# Stock Analysis with VBA (Refactored)

## Overview of Project
The purpose of this project was to analyze stock returns over a two year (2017 & 2018) period for 12 companies. Specifically, Total Daily Volumes and Starting/Ending Prices for each stock were collected. VBA script was developed, then refactored (edited to make more efficient), to parse the daily stock returns and to provide a clean formatted table that allows the reader to quickly determine performance of each stock involved.

## Analysis
Because the the volume of data (over 3,000 lines of returns for both years) and the nature of the analysis (finding the starting and ending prices for each stock), parsing the data using traditional Excel functions would be a cumbersome challenge to replicate year-over-year as the data set changes in scale. To overcome these challenges and to create a process that could be replicated for any year, the analysis uses VBA script to compile the data needed.

The major steps of the VBA script include:

#### 1) Creating the required arrays to store the data

Array of 12 for the different stocks in the analysis.
```    
    Dim tickers(0 To 11) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
```
3 arrays for each of the main points of analysis. 
 ```
    Dim tickerVolumes(0 To 11) As Long
    Dim tickerStartingPrices(0 To 11) As Single
    Dim tickerEndingPrices(0 To 11) As Single
```

#### 2) The main loop to compile the points of analysis.

Only the top part of the main loop shown:
```
        If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
```

#### 3) Loop to provide the compiled data to the table
```
    For i = 0 To 11
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = ((tickerEndingPrices(i) / tickerStartingPrices(i)) - 1)
```

Besides the the main data analysis, the goal of this project was to create an easily repeatable and efficient process. Taking existing code, we refoctored it (edited it) to make it more efficient. The main place we gained efficiency was in the primary loop, where a single loop was used instead of a nested loop. This can be see below:

#### Old Loop (Less Efficient)
```
For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets(yearValue).Activate
       For j = 2 To RowCount
       
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
```
#### Updated Loop (More Efficient)
```
Worksheets(yearValue).Activate
    For i = 2 To RowCount
    'tickerIndex = 0
    
        '3a) Increase volume for current ticker using the tickerIndex variable
        If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
        End If
```

The impact of this refactoring can be seen in how quickly the results occur. (See below)

### Analysis of 2017 Stocks
2017 was a good year for 11 of the 12 stocks invovled in the analysis. 11 of them saw a positive return from beginning of the year to the end. Further, several stocks saw over 100% gains in starting and ending prices over the year: DQ, ENPH, FSLR and SEDG.

We can also see this analysis ran in 0.14 seconds!

![VBA_Challenge_2017](https://user-images.githubusercontent.com/89284280/131058068-2754ff32-6332-40a2-8947-45a11e0272d7.PNG)

### Analysis of 2018 Stocks
2018 was a bad year for 10 of the 12 stocks invovled in the analysis, with these 10 seeing declines in prices from beginning of the year to the end. Specifically, 3 of the 4 stocks we saw in 2017 with the highest increases in price saw declines in 2018: DQ, FSLR and SEDG.

We can also see this analysis ran in 0.086 seconds!

![VBA_Challenge_2018](https://user-images.githubusercontent.com/89284280/131058134-58d3f998-6cd8-4860-8d5d-672809f5d488.PNG)


## Summary

#### What are the advantages or disadvantages of refactoring code?
The obvious benefit of refactoring code is making the code more efficient to run. This becomes more obvious as the data volume gets larger and larger. It would become very advantageous if you were running this same analysis on the entire S&P 500 versus just 12 stocks.

Additionally, refactoring your code forces the writer to think more aggressively about their code structure and methodology. Through iteration and testing you learn how to become more efficient. Hopefully this means the next time you have to perform a similar task that you begin with a mroe efficient code structure.

The only real disadvantage is the time is takes to go through the refactoring process. Depending on the scope of the project it may not require refactoring to accomplish your task.

#### How do these pros and cons apply to the refactoring the original VBA script?
Since the data set was relatively small (3,000 rows of data per year), the time it took for both the original and refactored script was very small (less than a second). And both methodologies got the same result. The question is then, is it worth it to refactor code that already works to gain 0.5s of time? Perhaps not. But, as the data set scales up, the value of the refactored code becomes more and more obvious.
