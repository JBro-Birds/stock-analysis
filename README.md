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
For the front-end user there are two macro buttons on the analysis worksheet, one for clearing data on the worksheet of a prior run and one for running the analysis for a given year.  The macro button asking the user for what year the analysis is for is an input box in which the user needs to enter a year.  The main part of the clear data script is `Cells.Clear` and for the the year selection input box script it is `yearValue = InputBox("What year would you like to run the analysis on?")`.  The year selection box is the first step in the `Sub AllStocksAnalysisRefactored()` refactored script.  After the user enters a year and clicks 'enter' a runtime timing variable `startTime = Timer` starts running as the starting point of the total process time that is shown at the end of the script; in addition the worksheet header in cell A1 is updated to include the year entered. 

A key component of the refractor script is defining the array `Dim tickers(12) As String` and initializing the 12 tickers; starting with the first ticker `tickers(0) = "AY"` and ending with the final ticker `tickers(11) = "VSLR"`.  Three output arrays are defined for the use in `for` loops and `if-then` statements in order for data to be stored as the script runs through the ticker index `tickerIndex = 0`  These three arrays are defined as follows:  `Dim tickerVolumes(12) As Long`, `Dim tickerStartingPrices(12) As Single`, `Dim tickerEndingPrices(12) As Single`.  These three arrays are initially set to "0" for the purpose of `for` loops and `if-then` statements.  

Next a `for` loop determines the dataset range `For i = 2 To RowCount` (with `Rowcount` defined as `RowCount = Cells(Rows.Count, "A").End(xlUp).Row`) and pulls the volume data by `tickerIndex` with the following formula:  `tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value`.  Next an `if-then` block determines the starting and ending rows for each `tickerIndex` and when to move on to the next `tickerIndex` in order to pull the `tickerStartingPrices(i)` and `tickerEndingPrices(i)`.  These formulas are as follows:  
`If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then tickerStartingPrices(tickerIndex) = Cells(i, 6).Value` 

`If Cells(i + 1, 1).Value <> tickers(tickerIndex) Cells(i, 1).Value = tickers(tickerIndex) Then tickerEndingPrices(tickerIndex) = Cells(i, 6).Value`.

`If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then tickerIndex = tickerIndex + 1`

After pulling the array data by `tickerIndex` a `for` loop is run to populate the defined and corresponding cells in the "All Stocks Analysis" worksheet. These values are for `tickers(i)`, `tickerVolume(i) and "Return" formulated as `tickerEndingPrices(i) / tickerStartingPrices(i) - 1`.

The script runs through a number of formatting statements and then the `endTime = Timer` is triggered.  The `endTime` less `startTime` is the total elasped time of the `Sub AllStocksAnalysisRefactored()` execution and a messagebox will appear to the end user stating the execution time.  `MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)`.

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

### Advantages and disadvantages of refactoring code in general
The advantages of refactoring code in general is that it helps make a program run faster, adds efficencies in storage and helps with potential debugging.  The disadvantages could be when much time is spent on refactoring for little gain in speed or efficiences or the refactoring code contains bugs.


### Advantages and disadvantages of the original and refactored VBA script
The advantage of the refactored VBA script is that it reduced the run time vs. the original script by ~80%.  Vice versa, a disadvantage of the refactored script is if this was the final project request from the customer a lot of time was spent on refactoring the code for what really amounted to not much execution time saved in relative terms when the original script would had sufficed.  Now if the dataset was increased with a substantial amount of additional tickers this statement probably would not hold true and the total process time saved would be worth the refactoring effort.  Another advantage for the refactored script is by using more arrays the data storage is more efficient compared to the original script.
