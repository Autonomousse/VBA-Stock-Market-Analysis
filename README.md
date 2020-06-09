# VBA-challenge
VBA - Stock market data analysis

## Goal
The goal of this assignment is to take stock data, stored in an Excel workbook, and create a summary table using VBA. We are also using VBA to highlight and format the summary table for a cleaner appearance. The code should loop through every worksheet and row and produce the following summary table items:

  - **Ticker**: the abbreviation of the stock name.
  
  - **Yearly Change**: the difference between the opening price at the beginning of a given year to the closing price at the end of that year.
  
    - The cell color should be green if postive and red if negative.
    
  - **Percent Change**: the percent change from opening price at the beginning of a given year to the closing price at the end of that year.
  
    - This column should be formatted as a percent with two decimals.
    
  - **Total Stock Volume**: the total stock volume summed up for any given year per stock.
  
  ## Challenge
  The challenge portion of this assignment was to create a second summary table that displays the following information along with the ticker and value:
  
  - **Greatest % Increase**
  
  - **Greatest % Decrease**
  
  - **Greatest Total Volume**
  
The second part of the challenge was to make it so the script could iterate through every worksheet in the workbook by running the VBA script once.
