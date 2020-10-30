# stock-analysis

## Overview of the Project

The stock analysis project shows the comparison between original and refactored VBA scripts. In this analysis, we have two codes that analyzes an entire dataset the same way. Based on the findings, it is evident that the refactored code is more efficient because it takes fewer steps, uses fewer memory, and takes less time. The performance of the original and refactored scripts were measured based on the run time. The project highlights the advantages and disadvantages of the original and refactored VBA scripts, and the performance of each script.


### Purpose

The purpose of the stock analysis is to edit and refactor the original VBA script. Although the existing VBA script works well, it takes some time to execute. There is another way to accomplish the task in the event a future user wanted to expand the dataset. By refactoring the code, we can loop through all the data one time as opposed to twice. Refactoring the code outputs the same result but in a timely manner. 

## Results

### Analysis of the Outcomes
 1. A `tickerIndex` variable was created, the dimension was defined as an integer, and set equal to zero. 
 2. Three output arrays were created and the dimensions were defined accordingly.
    - `tickerVolumes(11)` as Long
    - `tickerStartingPrices(11)` as Single
    - `tickerEndingPrices(11)` as Single 
3. A `for` loop was initialized to run through the tickers and the `tickerVolumes` was set to zero.
4. Another `for` loop was created to loop over all the rows in the spreadsheet.
5. Inside the `for` loop the first step is to calculate ticker volume or the sum of the volume for each ticker. 
6. An `if-then` statement was created to check if the current row is the first row with the selected `tickerIndex`. If the condition was met, then the `tickerStartingPrices` would be assigned the current closing price. 
        
        If Cells(i - 1, 1).Value <> Cells(i, 1) And Cells(i, 1).Value = tickers(tickerIndex) Then

            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                          
        End If
    This statement is checking if the current and previous row are the same the ticker. If the values the cells are not the same, then the program knows a new ticker has been selected and new starting price needs to be outputted.

7. An `if-then` statement was created to check if the current row is the last row with the selected ticker. 
        
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i + 1, 1).Value <> Cells(i, 1) Then

            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
    This statement is checking if the current and following row are the same. If values in the cell does not match the ticker and the row are of a different ticker, then the program knows that a new ticker has been selected and the ending price needs to be outputted.
8. An `if-then` statement was created to increase the `tickerIndex` if the next row's ticker does not match the previous row's ticker. If the conidtion in Step 7 has been met, the program will loop again with the next ticker.
9. A `for` loop was created to loop through the arrays `tickers`, `tickerVolumes`, `tickerStartingPrices` and `tickerEndingPrices`. The results were outputted to "Ticker", "Total Daily Volume", and "Return" columns in the spreadsheet. 

The following images represent the run time for years 2017 and 2018 using the refactored code: 

![](https://github.com/irenedepacina/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)


![](https://github.com/irenedepacina/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary 

### What are the advantages or disadvantages of refactoring code?

- Advantages:
    - The program is more efficient 
    - The program requires lesser memory
    - The program is easier to understand
    - The program runs faster 
 
- Disadvantages:
    - The program presented a lot of bugs
    - Time consuming

The following images shows the performance of the original and refactored script based on run time in 2017:

![](https://github.com/irenedepacina/stock-analysis/blob/main/Resources/VBA_original_2017.png)

![](https://github.com/irenedepacina/stock-analysis/blob/main/Resources/VBA_refactored_2017.png)

### How do these pros and cons apply to refactoring the original VBA script?

The refactored code is building on the original VBA script as opposed to recontructing a new code from scratch. The purpose of refactoring the original VBA script is to make the program advantageous. For the most part, the stock analysis demonstrates that the refactored code performed better than the original. The program was more efficient, required less time to run and was easier to understand in comparison to the original script. The existing VBA script consisted of a nested loop. Nested loops are easier to follow and conceptualize but such loops are not time efficient when executed. When applying refactoring, it was apparent that doing so introduced more bugs and required more time to code. 
 
