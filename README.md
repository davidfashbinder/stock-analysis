# VBA of Wall Street 

## Overview of Project
---
The purpose of this project is to refactor our original code used to analyze the stock 'DQ' and expand our analysis to the entire stock market.  


### Purpose
We will determine what benefits are inherent to using refactored code, along with expanding our analysis of each stock's total volume and its Return Value %.

### Results 

### 2017 vs. 2018

![2017](https://github.com/davidfashbinder/stock-analysis/blob/master/Stocks_2017.png)
![2018](https://github.com/davidfashbinder/stock-analysis/blob/master/Stocks_2018.png)

1. We can see that every stock, with the exception of ENPH and RUN posted negative returns in 2018.  In 2017, only TERP posted a negative return.  TERP and RUN's rate of return improved year-over-year from 2017 to 2018, while everyone else's dropped.  

![2017_Run_Time](https://github.com/davidfashbinder/stock-analysis/blob/master/VBA_Challenge_2017.png)
![2018_Run_Time](https://github.com/davidfashbinder/stock-analysis/blob/master/VBA_Challenge_2018.png)

2. As you can see in the screenshots above, the new refactored code runs very fast in both years.  This should give us confidence in our ability to use this code with larger data sets to compare more years, more stocks, or deploy data to more arrays; such as year-over-year performance.  

![Refactored_Code](https://github.com/davidfashbinder/stock-analysis/blob/master/Code1.png)

Some of the most key changes to the code can be seen above, where we added definitions to yearValue to allow any sheet to be activated from the input box that prompts when the code first runs; where the RowCount is automated; and where we create a tickerIndex to move away from hard-coding the ticker number.  This allows the code to run much more smoothly!

### Summary
1. Advantages of refactoring code include, but are not limited to:
-Saving time not having to re-write code that already worked
-Reducing the likelihood of coding errors since existing code works
-Being able to take a problem you solved from one data set and apply it to a different one

2. Disadvantages of refactoring code include, but are not limited to:
-Potentially long debugging processes as you adjust your code for the new data set or problem
-Ruining or deleting working code if you do not save a copy of your original code first
-Using code that is designed for small data sets on large data sets, increasing run times or preventing compilation.

In refactoring this script, I ran into several issues, primarily around iterators and subsets relating to each other in various For Loops and If Then statements.  While attempting to debug, I caused other issues, such as broken code where I had to force-close Excel and start over.  Since we were not hard-coding values, I also struggled with definining variables and iterators throughout the code.
