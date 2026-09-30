<h1 align="center">📊 Financial Data Analysis</h1>

<p align="center">
A Python-based financial analysis project leveraging real-world stock data to evaluate company performance
</p>


---
##  How It Works

 - Stock tickers are read from one or more labeled rows of `stocklist.txt` (for example `CD:` or `AB:`)
 - The program retrieves:
   - Company profile data
   - Historical stock prices
 - The data is processed to calculate:
   - Volatility (standard deviation)
   - Price trend (slope)
 - Results are saved into an Excel file with a summary of the findings

---

##  Setup & Usage

1. Install the dependencies:
   ```bash
   pip install requests openpyxl
   ```
2. Get a free API key from [Financial Modeling Prep](https://site.financialmodelingprep.com/) and set it as an environment variable:
   ```bash
   export FMP_API_KEY="your_key_here"      # macOS/Linux
   set FMP_API_KEY=your_key_here           # Windows
   ```
3. Run the program, optionally choosing which rows of `stocklist.txt` to analyze (default is `CD`):
   ```bash
   python project_3_start.py          # uses the CD row
   python project_3_start.py AB,CD    # combines several rows
   python project_3_start.py ALL      # uses every row
   ```

---

##  Key Features

-  Developed a Python-based financial analysis program using the Financial Modeling Prep API to retrieve and process real-world data for a configurable list of publicly traded companies  
-  Calculated volatility (standard deviation of closing prices) and price trend (average daily price change) over the most recent 30 trading days to compare company performance  
-  Built reusable Python functions to automate data retrieval, analysis, and reporting, producing an Excel report with a summary of the most and least volatile stocks and the strongest upward and downward trends  
-  Kept the API key out of the source code by reading it from an environment variable  

---

##  Output

The program generates an Excel file containing:

- **Summary** → Overall averages plus the most/least volatile companies and the strongest upward/downward trends  
- **Company Data** → Company name, sector, exchange, volatility, and trend slope for each ticker  
- **Stock Data** → Daily open and close prices for the last 30 trading days  

---

##  Technologies Used

- Python  
- `requests` (API calls)  
- `openpyxl` (Excel file generation)  
- `statistics` (data analysis)  

---

##  Conclusion

I worked on this project to demonstrate my ability to:  
 - Pull and use data from real-world financial APIs  
 - Work with and analyze large datasets  
 - Apply statistical methods to make sense of financial data  
 - Build reusable and scalable Python functions  

---

<p align="center">
Built for financial data exploration and analysis 📈
</p>
