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

