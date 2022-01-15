# stocks-analysis

## Overview of Project
Steve is a recent finance graduate who is helping his parents, who are also his first clients, with stock investments.  They are interested in investing in green energy, and have already bought stock in one such company, "DQ".  However, they invested in "DQ" without doing much research into its returns.  

### Purpose
Steve wants to look into "DQ"s returns over 2 previous years, as well as those of several other green energy companies that are traded, to help diversify his parents' investments. This was done by writing a VBA script that compiles each company's total daily volume as well as calculates the percent of return for each year.  

## Results


## Summary
### Advantages/Disadvantages of Refactoring Code
There are several advantages and disadvantages of refactoring code.  Refactoring makes the code less complicated, which can make it easier for someone to read adn understand.  It also makes for "cleaner" code.  While the function doesn't change, refactored code should run  more efficiently than previous.  It is also easier to debug when code is clean and understandable, as well as making it easier to maintain and update the code as needed. However, refactoring can introduce bugs into the code.  Also, while changes to the code may impact efficiency (such as an increase in run speed), this may not produce enough of a benefit compared to the amount of time and effort spent refactoring.  

### Advantages/Disadvantages of Refactoring VBA Challenge Script
Refactoring the original VBA script did increase its efficiency.  The code did run more quickly, as evidenced by the differences in our timer messages pre- and post-refactoring. The refactoring did potentially make the code less complicated, although for an inexperienced programmer (like myself), the introduction of the 3 output arrays for volumes, starting & ending prices versus the original individual variables created a greater sense of confusion in creating the loops.  In this sense, the marginal increase in speed after refactoring may not have been worth the time spent creating the arrays and refactoring the loops.  It would seem that refactoring is likely best left to experienced programmers with a strong knowledge of patterns and coding techniques.    
