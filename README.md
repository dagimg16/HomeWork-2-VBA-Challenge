# Module-2-VBA-Challenge

## Quarterly Stock Analysis (VBA Script)

## About
This project is part of my Data Science course at SMU, where I’m learning how to apply programming to analyze and process data. 
The purpose of this VBA script is to analyze stock data on a quarterly basis and calculate some important metrics:
- **Ticker**: The stock identifier.
- **Quarterly Change**: The difference in stock price from the start to the end of the quarter.
- **Percent Change**: The percent increase or decrease during the quarter.
- **Total Stock Volume**: The total number of shares traded during the quarter.
- **Greatest % Increase**: The greatest percent increase in the quarter
- **Greatest % Decrease**: The greatest percent decrease in the quarter
- **Greatest Total Volume**: The greatest total volume in the quarter

## What You Need
- Microsoft Excel with VBA enabled.
- A dataset with the following columns:
  - **Ticker symbol** in Column A 
  - **Dates** In Column B
  - **Opening prices** in Column C
  - **High Prices** in Column D
  - **Low Prices** in column E
  - **Closing prices** in column F
  - **Trading volumes** in Column G
- Each Sheet should only contain data for a signle quarter
    
## How to Use
1. Open your Excel file and press `Alt + F11` to access the VBA editor.
2. Insert a new module by right-clicking on the workbook and selecting **Insert > Module**.
3. Copy and paste the script into the module.
4. Save your file as a macro-enabled workbook (`.xlsm`).
5. To run the script:
   - Press `Alt + F8` in Excel.
   - Select the macro (e.g., `QuarterlyStockAnalysis`) and click **Run**.
6. The output will be displayed in the same sheet starting from column I.

## Example
- Please refer to the screenshot provided for an example of the input data structure and expected output.

## Learning Experience
This project has been a great way to:
- Practice VBA for data manipulation and analysis.
- Work with financial datasets and derive meaningful insights.
- Develop coding skills in a structured environment as part of my course at SMU.

## Acknowledgments
- Thanks to SMU and my instructors for guiding me through this learning journey!
- And a big thank you to everyone checking out this project.

## License
Feel free to use and modify this script for your own learning.
