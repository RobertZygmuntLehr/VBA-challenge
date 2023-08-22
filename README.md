# VBA of Wall Street Stocks

## Purpose
Stock market data is analyzed using Excel VBA scripting. The provided data contains stock information for various companies over the course of a year. The objective is to create a script that calculates and displays the yearly change, percentage change, and total stock volume for each stock. Additionally, the script identifies the stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume.

## The Script
1. __Retrieval of Data:__ The script loops through one year of stock data and reads/stores the ticker symbol, volume of stock, open price, and close price for each row.
2. __Column Creation:__ The script creates columns for the ticker symbol, total stock volume, yearly change, and percentage change.
3. __Conditional Formatting:__ Conditional formatting is applied in the yearly change column to highlight positive and negative changes appropriately.
4. __Calculated Values:__ The script calculates the greatest percentage increase, greatest percentage decrease, and greatest total volume for the stocks.
5. __Looping Across Worksheet:__ The VBA script successfully runs on all sheets in the workbook.

### Tools and Techniques Used
- VBA scripting
- Looping through data
- Calculating metrics (yearly change, percentage change, total volume)
- Conditional formatting (highlighting positive/negative change)
- Worksheet navigation and automation
- Modularizing code for reusability