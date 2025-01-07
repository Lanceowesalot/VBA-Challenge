# VBA Challenge 2 - Stock Market Analysis

 
Instructions:

Create a script that loops through all the stocks for each quarter and outputs the following information:

•	The ticker symbol

•	Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.

•	The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.

•	Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".

•	Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every quarter) at once.


About the Macro:

This VBA macro, Sub StockTicker(), processes stock market data from multiple worksheets in an Excel workbook. The macro analyzes stock performance by calculating metrics such as yearly change, percent change, and total stock volume for each ticker symbol. It also identifies the greatest percentage increase, greatest percentage decrease, and greatest total volume across all the tickers.
What the Macro Does:
1.	Iterates Through Worksheets:
o	The macro loops through all worksheets in the workbook.

2.	Adds Column and Row Headers:
Inserts headers for analysis in columns I to L:
o	"Ticker"
o	"Yearly Change"
o	"Percent Change"
o	"Total Stock Volume"
	Adds summary headers in columns O to Q for the greatest values:
o	"Greatest % Increase"
o	"Greatest % Decrease"
o	"Greatest Total Volume"

3.	Initialize Variables:
   
    o	Tracks key values like ticker, begin_price, end_price, change_amount, percent_change, and cumulative stock_volume.

4.	Process Stock Data:
   
    o	Loops through each row of data (Column A for tickers and Column G for stock volume).

  	o	Identifies when a ticker changes or when the last row is reached.

  	o	Calculates:
  	
        	Yearly Change (end_price - begin_price).
  	
        	Percent Change ((change_amount / begin_price)).
  	
    o	Populates columns:
  	
        	I: Ticker name.
  	
        	J: Yearly change.
  	
        	K: Percent change.
  	
        	L: Total stock volume.

5.	Highlight Yearly Change:
   
    o	Sets cell background color in column J based on the sign of change_amount:

        	Red: Negative change.
        
        	Green: Positive change.
        
        	White: No change.

6.	Track Greatest Values:
   
    o	Identifies the ticker with:
  	
        	Greatest percentage increase.
  	
        	Greatest percentage decrease.
  	
        	Greatest stock volume.
  	
    o	Outputs these values in columns P and Q.

7.	Adjust Column Widths:
   
    o	Automatically adjusts column widths to fit the content.

Usage:

This macro is to be used for analyzing historical stock data organized by ticker. It processes data for each worksheet, making it suitable for workbooks with multiple sheets, where each sheet contains stock trading data for different periods.

Reference:

Chat GPT for debugger and code suggestions

Creations and use of headers -

    o Stackoverflow. https://stackoverflow.com/questions/31540366/vba-excel-how-to-insert-a-predetermined-row-of-headers

    o Stackoverflow https://stackoverflow.com/questions/39705104/insert-text-array-as-headers-vba-excel

Class instruction

Format help and some code-

    o GitHub – user Rooprg https://github.com/rooprg/vba_challenge_stock_prices/blob/main/VBA_Code_All_Worksheets_rroop_11Apr2024.vbs
