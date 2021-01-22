# Green Stock Analysis
## Project Overview
In this project I was able to help Steve analyze data in order for him to have a better understanding of Green Energy stocks to help his parents future investments. I utilized Visual Basic for Applications, VBA, through excel in order to present the activity in different stocks throughout the years 2017 and 2018. The stock activity was measured by obtaining the daily volume or number of shares as well as the percentage of annual return per stock. Initially, I focused my analysis on the parents preferred stock, DQ, and then I was able to do an analysis on the other stocks in the data to present which stocks were a better investment for Steve’s parents moving forward. 
#### Purpose
The purpose of this analysis was to see which Green Energy stocks were going to be the best investment for Steve’s parents. This was completed by running an analysis on the preferred stock, DQ, and then for the other 11 stocks. I completed my initial analysis for all stocks and then was purposely challenged to refactor my code to make it run more efficiently and readily throughout the data. The overall scheme of this project was to provide comprehensible stock data in different efficient ways. 
## Results
#### Analysis
The refactoring of my original script was done by changing my variables, adding a tciker index, changing output arrays, and changing my for loops into one consistent for loop to run the code more efficiently. 

Original Code

```
3a) Initialize variables for starting price and ending price
    Dim startingPrice As Single
    Dim endingPrice As Single

'3b) Activate data worksheet
Worksheets(yearValue).Activate
    
'3c) Get the number of rows to loop over
    'rowCount taken from stackoverflow
    'rowCount taken from module hint
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
'4) Loop through tickers
     For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0 
'4) Loop through tickers
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
    
'5b) get starting price for current ticker
    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        startingPrice = Cells(j, 6).Value
    End If
        
'5c) get ending price for current ticker
    If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        endingPrice = Cells(j, 6).Value
    End If
    
    Next j
     
'6) Output data for current ticker
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
    Next i
    
```

Refactored Code
```
    '1a) Create a ticker Index
        Dim tickerIndex As Single
        tickerIndex = 0

    '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
        
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
        tickerVolumes(i) = 0
        
        Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        
        '3a) Increase volume for current ticker
        'taken from challenge hint
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
 
            '3d Increase the tickerIndex.
         tickerIndex = tickerIndex + 1
                
         End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    Worksheets("All Stocks Analysis").Activate
    For i = 0 To 11
    tickerIndex = i
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

```

    

#### Stock Performance
The overall stock performance differed in both the years 2017 and 2018.  In 2017, 11 of the 12 Green Energy stocks had a positive annual return, meaning that their value from the beginning of year to the end of the year had a positive return. Meanwhile for the year 2018, there was a negative and drop in the percent of annual return in 10 out of the 12 stocks. The overall return on green stocks dropped in 2018. This analysis was done by obtaining the totalVolume and return for each stock. As well as adding formatting and conditional formatting functions in VBA to read the data more efficiently. 

#### Run Time Performance
There was a major difference when it came to execution times between the original and refactored VBA script. The first time I ran my original script I got a ridiculous number for both years, partially I believe my excel document was lagging.
When I refactored my code it cut down my run time by over 67,000 seconds for both years. 

Original code run time


![](https://github.com/allison-chavez/stock-analysis/blob/main/Resources-VBA/Screen%20Shot%202021-01-21%20at%206.52.30%20PM.png)
![](https://github.com/allison-chavez/stock-analysis/blob/main/Resources-VBA/Screen%20Shot%202021-01-21%20at%206.52.48%20PM.png)

Refactored code run time

![](https://github.com/allison-chavez/stock-analysis/blob/main/Resources-VBA/Screen%20Shot%202021-01-21%20at%207.26.24%20PM.png)
![](https://github.com/allison-chavez/stock-analysis/blob/main/Resources-VBA/Screen%20Shot%202021-01-21%20at%207.26.46%20PM.png)

## Summary
### Refactoring general code 
There are differing advantages to refactoring code, but overall it cuts down run time and allows analysis to be done more efficiently if done properly. The purpose of refactoring code is to improve the code that has been established, meaning to clean up the code and figure out a way to get the same result with less steps or in a way that is more linear when following and understanding. Some of the disadvantages I faced was trying to find more efficient or logical ways to refactor the code. It takes time to take the code you have already completed and try to manipulate it to make it run properly and efficiently. I believe there is room to make mistakes to the code and data results that have already been achieved with the original code.

### Refactoring the original VBA script
The advantage of refactoring the original script was that the original code could be used as a template for the refactored code. Obtaining a strong knowledge of the original code could be carried over when refactoring the code. The disadvantage for me was really understanding what the refactored code was asking for. In example, changing the names of the variables and trying to take the knowledge already engraved in my head and trying to modify it to get the same result. I think the major disadvantage is that there is a lot of room to get confused when refactoring my original VBA script when missing small points. 
