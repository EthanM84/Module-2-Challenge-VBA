Stock Analysis: 
This VBA macro, named StockAnalysis, is designed to analyze stock data for the years 2018, 2019, and 2020 in an Excel workbook. The macro performs the following tasks:

Loops through each year in the array of years ("2018", "2019", "2020").
Sets the active worksheet to the corresponding worksheet for the current year.
Calculates various financial metrics for each stock ticker in the data, including yearly change, percent change, and total volume.
Outputs the results to a summary table in the worksheet.
Applies conditional formatting to the yearly change cell to highlight positive changes in green and negative changes in red.
Identifies the ticker with the greatest percent increase, greatest percent decrease, and greatest total volume, and outputs the results to the summary table.
Formats the percent increase and percent decrease cells as percentages.
Usage: 
Open an Excel workbook containing stock data for the years 2018, 2019, and 2020.
Press Alt + F11 to open the Visual Basic Editor in Excel.
Insert a new module or open an existing module.
Copy and paste the StockAnalysis macro into the module.
Close the Visual Basic Editor.
Run the macro by going to the Excel worksheet with the stock data, and pressing Alt + F8 to open the Macro dialog box. Select "StockAnalysis" from the list of macros and click "Run".
The macro will analyze the stock data for each year and output the results to a summary table in each worksheet.
Note: Make sure to enable macros in your Excel settings before running the macro.

Requirements: 
This macro requires the following:

Microsoft Excel (version 2010 or later)
Stock data for the years 2018, 2019, and 2020 in separate worksheets in the Excel workbook.
Stock data should be organized with the following columns: ticker symbol (in column A), date (in column B), opening price (in column C), closing price (in column F), and volume (in column G).
Output: 
The macro will output the following results to the summary table in each worksheet:

Ticker symbol in column I
Yearly change in column J
Percent change in column K (formatted as a percentage)
Total volume in column L
Ticker with the greatest percent increase, greatest percent decrease, and greatest total volume in cells O2, O3, and O4 respectively
Corresponding values for the greatest percent increase, greatest percent decrease, and greatest total volume in cells Q2, Q3, and Q4 respectively
The macro will also apply conditional formatting to the yearly change cell in column J to highlight positive changes in green and negative changes in red.

Please ensure that you have a backup of your data before running the macro, as it may modify the contents of your workbook.
