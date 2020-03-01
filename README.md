# VBA Script

Created a script that will loop through all the stocks data in multiple sheets in a single Excel file. Each sheet included data for multiple stocks for that given year. Three years worth of data were analyzed. 

The different sheets ranged from 700,000 to 800,000 rows of data. Seven columns were available for analysis for each year: 
- ticker (ticker symbol)
- date	
- open (opining price)
- high (highlest price of day)
- low (lowest price of day)
- close (closing price)
- vol (volume)

The analysis generated the following information:
- The ticker symbol.
- Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
- The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
- The total stock volume of the stock.

The output was also conditionally formatted to automatically highlight positive change in green and negative change in red.
Output also included identification of the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".

Please note that the original files analyzed were in Excel format and were too big to upload to GitHub. As such, only the VBA code was uploaded to the repo.
