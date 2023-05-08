# Stock Market Analysis VBA

CODE SOURCE:

This VBA script was created by Nestor Baptista

OVERVIEW:

This VBA code is designed to analyze stock market data in multiple worksheets within an Excel workbook. The macro loops through each worksheet, calculating the yearly change, percentage change, and total stock volume for each stockID for each year. it then summarizes the data into a first table and highlights positive and negative changes in green and red, respectively. Additionally, the macro identifies the stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume for each worksheet and summarizes it into a second table.

LIMITATIONS:

The code has a few limitations:
1. It assumes that the stock data is sorted by ticker and date in ascending order.
2. It assumes that there are no missing or empty cells in the data.
3. It assumes that the data is in the same format as the data provided in the repository.
4. It may not work for very larger datasets due to memory constraints.
