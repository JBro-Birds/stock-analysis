# Stock Analysis - Refactored for Faster Runtime

## Overview of Project
The customer liked the original stock analysis workbook provided with the front-end ability to click a button in order to produce an analysis of an entire dataset. 
The customer now wants to expand the dataset to include the entire stock market over the last few years.  This project is to determine if the code written for the
original stock analysis project can be refactored to run a much larger dataset more efficiently in regards to execution runtime. 

### Purpose
The customer wants to expand the research to cover the entire stock market.  Prior coding was sufficient to execute the smaller dataset for runtime-purposes, but
with a much larger dataset code refactoring needs to be performed to determine if the executed runtime can be improved.  If refactoring is successful the runtime will run faster and the customer will remain happy. 

## Results

### Coding
For the front-end user there are two macro buttons on the analysis worksheet; one for clearing data on the worksheet of a prior run and one for running the analysis for a given year.  The macro button asking the user for what year the analysis is for is an input box in which the user needs to enter a year.  The main part of the clear data script is `Cells.Clear` and for the the year selection input box script it is `yearValue = InputBox("What year would you like to run the analysis on?")`.  After the user enters a year and clicks 'enter' a runtime timing variable `startTime = Timer` starts running as the starting point of the total process time that is shown at the end of the script; in addition the worksheet header in cell A1 is updated to include the year entered. 



### Original Script Runtimes
The original script runtimes for 2017 and 2018 are as follows:

#### 2017
![2017 Stocks - Original Script](https://raw.githubusercontent.com/JBro-Birds/stock-analysis/master/Resources/VBA_Challenge_2017_OriginalScript.png)

#### 2018
![2018 Stocks - Original Script](https://raw.githubusercontent.com/JBro-Birds/stock-analysis/master/Resources/VBA_Challenge_2018_OriginalScript.png)






### Refactored Script Runtimes

#### 2017
![2017 Stocks - Refactored Script](https://raw.githubusercontent.com/JBro-Birds/stock-analysis/master/Resources/VBA_Challenge_2017_RefactoredScript.png)

#### 2018
![2018 Stocks - Refactored Script](https://raw.githubusercontent.com/JBro-Birds/stock-analysis/master/Resources/VBA_Challenge_2018_RefactoredScript.png)


## Summary

###

