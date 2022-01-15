# stocks-analysis

## Overview of Project
Steve is a recent finance graduate who is helping his parents, who are also his first clients, with stock investments.  They are interested in investing in green energy, and have already bought stock in one such company, "DQ".  However, they invested in "DQ" without doing much research into its returns.  

### Purpose
Steve wants to look into "DQ"s returns over 2 previous years, as well as those of several other green energy companies that are traded, to help diversify his parents' investments. This was done by writing a VBA script that compiles each company's total daily volume as well as calculates the percent of return for each year.  

## Results
An analysis of all stocks in the years 2017 and 2018 showed that overall, all but 2 stocks (TERP and RUN) had a significantly higher percentage of return in 2017, when compared to 2018's values.  While all but one stock (TERP) had positive returns in 2017, only 2 (ENPH and RUN) had positive returns in 2018.  ![2017 Green Stocks Performance](https://github.com/MischievousBadger/stock-analysis/blob/f99debb7be0f8effd18ce61b290ac355b4e45114/Resources/stocks_performance_2017.PNG)  ![2018 Green Stocks Performance](https://github.com/MischievousBadger/stock-analysis/blob/f99debb7be0f8effd18ce61b290ac355b4e45114/Resources/stocks_performance_2018.PNG) 
Steve's parents' stock, "DQ", while having an increase in total daily volume of 72,077,700 between 2017 and 2018, had a difference of -262% between the return values in 2017 and 2018.  This would be a significant argument to bolster Steve's desire to diversify his parent's investments.  

The original code to analyze all stocks was refactored for efficiency, using arrays for ticker volumes, starting prices and ending prices in the refactored code versus coding these as variables in the original. A ticker index was also added to the refactored code.  This resulted in an increase in execution time for the refactored code over the original. ![2017 Original Code](https://github.com/MischievousBadger/stock-analysis/blob/0dc35950e61e7ca518c2b24427b3d45a9e4e89ed/Resources/original_script_2017.PNG) ![2017 Refactored Code](https://github.com/MischievousBadger/stock-analysis/blob/0dc35950e61e7ca518c2b24427b3d45a9e4e89ed/Resources/VBA_Challenge_2017.PNG) 
![2018 Original Code](https://github.com/MischievousBadger/stock-analysis/blob/0dc35950e61e7ca518c2b24427b3d45a9e4e89ed/Resources/original_script_2018.PNG) ![2018 Refactored Code](https://github.com/MischievousBadger/stock-analysis/blob/0dc35950e61e7ca518c2b24427b3d45a9e4e89ed/Resources/VBA_Challenge_2018.PNG)


## Summary
### Advantages/Disadvantages of Refactoring Code
There are several advantages and disadvantages of refactoring code.  Refactoring makes the code less complicated, which can make it easier for someone to read adn understand.  It also makes for "cleaner" code.  While the function doesn't change, refactored code should run  more efficiently than previous.  It is also easier to debug when code is clean and understandable, as well as making it easier to maintain and update the code as needed. However, refactoring can introduce bugs into the code.  Also, while changes to the code may impact efficiency (such as an increase in run speed), this may not produce enough of a benefit compared to the amount of time and effort spent refactoring.  

### Advantages/Disadvantages of Refactoring VBA Challenge Script
Refactoring the original VBA script did increase its efficiency.  The code did run more quickly, as evidenced by the differences in our timer messages pre- and post-refactoring. The refactoring did potentially make the code less complicated, although for an inexperienced programmer (like myself), the introduction of the 3 output arrays for volumes, starting & ending prices versus the original individual variables created a greater sense of confusion in creating the loops.  In this sense, the marginal increase in speed after refactoring may not have been worth the time spent creating the arrays and refactoring the loops.  It would seem that refactoring is likely best left to experienced programmers with a strong knowledge of patterns and coding techniques.    
