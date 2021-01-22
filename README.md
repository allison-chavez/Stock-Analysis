# Green Stock Analysis
## Project Overview
In this project I was able to help Steve analyze data in order for him to have a better understanding of Green Energy stocks to help his parents future investments. I utilized Visual Basic for Applications, VBA, through excel in order to present the activity in different stocks throughout the years 2017 and 2018. The stock activity was measured by obtaining the daily volume or number of shares as well as the percentage of annual return per stock. Initially, I focused my analysis on the parents preferred stock, DQ, and then I was able to do an analysis on the other stocks in the data to present which stocks were a better investment for Steve’s parents moving forward. 
#### Purpose
The purpose of this analysis was to see which Green Energy stocks were going to be the best investment for Steve’s parents. This was completed by running an analysis on the preferred stock, DQ, and then for the other 11 stocks. I completed my initial analysis for all stocks and then was purposely challenged to refactor my code to make it run more efficiently and readily throughout the data. The overall scheme of this project was to provide comprehensible stock data in different efficient ways. 
## Results
#### Analysis
The refactoring of my original script was done by changing my variables, adding a tciker index, changing output arrays, and changing my for loops into one consistent for loop to run the code more efficiently. 

#### Stock Performance
The overall stock performance differed in both the years 2017 and 2018.  In 2017, 11 of the 12 Green Energy stocks had a positive annual return, meaning that their value from the beginning of year to the end of the year had a positive return. Meanwhile for the year 2018, there was a negative and drop in the percent of annual return in 10 out of the 12 stocks. The overall return on green stocks dropped in 2018. This analysis was done by obtaining the totalVolume and return for each stock in VBA using the code:
#### Run Time Performance
There was a major difference when it came to execution times between the original and refactored VBA script. The first time I ran my original script I got a ridiculous number for both years, partially I believe my excel document was lagging.
When I refactored my code it cut down my run time by over 67,000 seconds for both years. 

![](https://github.com/allison-chavez/stock-analysis/blob/main/Resources-VBA/Screen%20Shot%202021-01-21%20at%206.52.30%20PM.png)
## Summary
### Refactoring general code 
There are differing advantages to refactoring code, but overall it cuts down run time and allows analysis to be done more efficiently if done properly. The purpose of refactoring code is to improve the code that has been established, meaning to clean up the code and figure out a way to get the same result with less steps or in a way that is more linear when following and understanding. Some of the disadvantages I faced was trying to find more efficient or logical ways to refactor the code. It takes time to take the code you have already completed and try to manipulate it to make it run properly and efficiently. I believe there is room to make mistakes to the code and data results that have already been achieved with the original code.

### Refactoring the original VBA script
The advantage of refactoring the original script was that the original code could be used as a template for the refactored code. Obtaining a strong knowledge of the original code could be carried over when refactoring the code. The disadvantage for me was really understanding what the refactored code was asking for. In example, changing the names of the variables and trying to take the knowledge already engraved in my head and trying to modify it to get the same result. I think the major disadvantage is that there is a lot of room to get confused when refactoring my original VBA script when missing small points. 
