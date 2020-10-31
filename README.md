# Refactor VBA Code
## Overview of Project
### Purpose
We create a solution for Steve to show his parents the daily trade volume and yearly return for each stock for 2017 and 2018.  We started with an excel file that contained stock data for "Green Stocks" which included the ticker, date, opening price, high and low for the day, closing price, adjusted closing price and the volume by day. 
To plan for the future the current solution needed to be improved to be more efficient so it can process any number of stocks, for any number of years in a shortest time possible. 

## Results
### Stock Comparison
After analysing the results for all the stocks in 2017 and 2018 we conclude that 2017 outperformed 2018.  
We can see in the 2017 report that the yearly return for each of the stocks was significantly higher compared to 2018.  The Daily return results are in green show a positive return for all but 1 stock. If you look at the 2018 report it shows that all but 2 stocks are red and had a negative return. 
The DQ stock that Steve’s parents were interested in and it has the biggest decrease in return from 2017 to 2018, going from 199.4% down to -62.6%.  
<img src="https://github.com/andralobo/Module2-Challenge/blob/main/VBA_Challenge_2017.png?raw=true">
<img src="https://github.com/andralobo/Module2-Challenge/blob/main/VBA_Challenge_2018.png?raw=true">
### Speed Comparison 
The old script completed in 0.83 seconds and the new refactored script ran at 0.22 seconds.  That means the new script ran 0.61 faster and provided the same results. Refactoring the code definitely improved the speed of processing and in the future, it will be scalable and be able to accommodate more data and run more efficiently.

## Summary
### Refactoring code - Advantages or Disadvantages
There are many advantages to refactoring code and by restructuring and optimizing you can reduce redundant code while keeping the same functionality.  Refactoring helps understand the code better and allows time to make better decision on design, structure and planning for future growth.  
The disadvantages of refactoring code are that it is time consuming and can be expensive.
 ### Pros and cons to refactoring the original VBA script
The advantages to refactoring the script we created was that I now understand the code better, I was able to optimize the code to make it run faster and most importantly the script is now scalable and can accommodate more data.
The disadvantage of refactoring the script was that it was time consuming and this is all.  In this case the advantages outweigh the only disadvantage I could find.   
