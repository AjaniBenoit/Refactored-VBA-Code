# Refactored VBA Code 

Analysis of refactored VBA code for Green Stocks excel file. 

## Overview of Project 

Visual Basic for Applications (VBA) was used to automate the analysis of the Green Stocks excel [VBA_Challenge.xlsm]( https://github.com/AjaniBenoit/Refactored-VBA-Code/blob/main/VBA_Challenge.xlsm) file to evaluate stock performance. The original VBA code allowed the user to input either the year 2017 or 2018 to calculate daily volume and yearly return for each stock.  To analyze larger dataset of stocks the VBA code was refactored to make it more efficient. 

### Purpose 

The purpose of this report is to outline the both the original and refactored coded; and to determine if the refactored code is more efficient. 

## Results 

### Original VBA Code.

The original code creates a subroutine called “yearValueAnalysis()”. This sub routine allows the user to enter the year they wish to run the analysis on. The Code then populates the year the report is being executed for along with the headers for the rows in the “All Stocks Analysis” worksheet. The Code then creates an array for the Tickers and creates two new variables for the starting price and ending price. ![Original_VBA_CODE_1.png]( https://github.com/AjaniBenoit/Refactored-VBA-Code/blob/main/Original_VBA_CODE_1.png)

The code then activates the worksheet for the year in which the analysis is being completed and counts then number of rows that the code would be looped through. ![Original_VBA_CODE_2.png]( https://github.com/AjaniBenoit/Refactored-VBA-Code/blob/main/Original_VBA_CODE_2.png) 

The Code then uses nested loops and conditional if-then statements to loop through the tickers and calculate the volume for the tickers, starting price of the tickers and ending price of the ticker. The code then outputs data for each ticker on the “All Stocks Analysis” worksheet. The subroutine is then ended
![Original_VBA_CODE_3.png]( https://github.com/AjaniBenoit/Refactored-VBA-Code/blob/main/Original_VBA_CODE_3.png)

### Refactored VBA Code

The refactored code creates a subroutine called “AllStockAnalysisRefactored()”. This subroutine allows the user to enter the year they wish to run the analysis on. The Code then populates the year the report is being executed for along with the headers for the rows in the “All Stocks Analysis” worksheet. The Code then creates an array for the Tickers. The code then activates the worksheet for the year in which the analysis is being completed and counts then number of rows for which the code would be looped through. ![Refactored_VBA_CODE_1.png]( https://github.com/AjaniBenoit/Refactored-VBA-Code/blob/main/Refactored_VBA_CODE_1.png) 

The refactored code then creates an new variable called “tickerIndex”. The ticker index is set to zero. The Code then creates three new arrays called “tickerVolumes”, “tickerStartingPrices”, and “tickerEndingPrices”. A loop is then utilized to set the tickerVolumes to zero. The loop is then closed. Another loop is then utilized to loop over all the rows in the spreadsheet. The tickerindex variable is then looped through the tickerVolumes Array to calculate the stocks volume. Conditional if-then statements are then used to loop the tickerIndex variable through the tickerStartingPrice Array and then the tickerEndingPrice Array to calculate the starting and ending price of each ticker. The loop is then closed. No nested loops are utilized in the refactored code.  ![Refactored_VBA_Code_2](https://github.com/AjaniBenoit/Refactored-VBA-Code/blob/main/Refactored_VBA_CODE_2.png)

The refactored code then initiates another loop through the arrays to output the tickers, total daily volumes and return. The loop is then closed. Another loop is then initiated and if-then conditional statements used to format the worksheet.  ![Refactored_VBA_CODE_3]( https://github.com/AjaniBenoit/Refactored-VBA-Code/blob/main/Refactored_VBA_CODE_3.png)

