# Deliverable Project - Automated Excel & Report Generator




## Overview
This project implements a sophisticated financial analysis tool that generates comprehensive reports for portfolio risk assessment. The tool performs several key functions:

### Key Features
1. **Data Collection**
   - Fetches historical price data for user-specified stock tickers using Yahoo Finance API
   - Automatically includes ORCL with 394 fixed shares
   - Supports flexible date ranges for analysis

2. **Risk Analysis**
   - Calculates daily returns and portfolio performance metrics
   - Computes Value at Risk (VaR) using multiple approaches:
     * 1-day historical VaR per stock
     * 10-day portfolio historical VaR
     * 10-day model-based VaR (normal distribution assumption)
   - Determines 10-day Conditional Value at Risk (CVaR)
   - Generates protective option strategies and payoff diagrams

3. **Portfolio Management**
   - Handles $1,000,000 portfolio allocation
   - Maintains fixed ORCL position (394 shares)
   - Optimizes remaining cash distribution among other stocks
   - Tracks cash positions and used capital

### Outputs
1. **Excel Workbook** (`Portfolio_VaR_Workbook.xlsx`)
   - Historical prices sheet
   - Returns calculations
   - VaR analysis (1-day and 10-day)
   - CVaR computations
   - Portfolio positions and cash allocation

2. **Word Report** (`Client_Report.docx`)
   - Executive summary
   - Portfolio construction details
   - VaR & CVaR analysis results
   - Options hedging recommendations
   - Currency swap design
   - Appendix with charts and figures

3. **Option Strategy Visualizations** (in `figures/` directory)
   - Protective put diagrams
   - ORCL straddle payoff charts
   - ORCL strangle strategy visualization

## Technical Requirements
- Python 3.9 or higher
- Required packages:
  * `yfinance`: Historical market data retrieval
  * `pandas`: Data manipulation and analysis
  * `numpy`: Numerical computations
  * `matplotlib`: Visualization and charts
  * `openpyxl`: Excel file generation
  * `python-docx`: Word document creation

## Installation
1. Install Python 3.9+ and pip
2. (Recommended) Create a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # or venv\Scripts\activate on Windows
   pip install -r requirements.txt
   ```

## Usage
### Basic Run
```bash
python generate_report.py --tickers AAPL MSFT NVDA --start 2024-11-08 --end 2025-11-08 --output ./output
```

### Advanced Options
1. **Custom Portfolio Analysis**
   ```bash
   python generate_report.py --tickers AAPL MSFT NVDA GOOGL META --start 2024-11-08 --end 2025-11-08 --output ./custom_analysis
   ```

2. **Longer Historical Period**
   ```bash
   python generate_report.py --tickers AAPL MSFT --start 2023-11-08 --end 2025-11-08 --output ./long_term
   ```

### Code Structure
- **Data Collection**: 
  * `download_data()`: Fetches historical price data using yfinance API
  * Handles data validation and error checking

- **Risk Calculations**:
  * `compute_returns()`: Calculates daily returns from price data
  * `historical_var()`: Computes historical VaR using percentile method
  * `model_var()`: Implements parametric VaR assuming normal distribution
  * `expected_shortfall()`: Calculates CVaR for tail risk assessment
  * `scale_var_1d_to_nd()`: Scales 1-day VaR to n-day VaR

- **Portfolio Management**:
  * `portfolio_value_from_weights()`: Calculates position sizes and allocations
  * `save_to_excel()`: Generates formatted Excel workbook with all calculations
  * Handles the fixed ORCL position and redistributes remaining capital

- **Options Analysis**:
  * `make_payoff_chart()`: Creates visualization of option strategies
  * Implements protective puts, straddles, and strangles
  * Visualizes risk/reward profiles

- **Report Generation**:
  * `generate_report_docx()`: Produces comprehensive Word report
  * Includes executive summary, analysis, and recommendations
  * Automatically integrates figures and charts

## Notes and Limitations
- The script requires active internet connection for data retrieval
- ORCL position (394 shares) is fixed and cannot be modified
- Option premiums are placeholder values; extend the code to fetch live data
- MarketWatch portfolio screenshots must be added manually to the report
- Default portfolio size is set to $1,000,000

## Troubleshooting
1. **Data Download Issues**
   - Ensure internet connectivity
   - Verify ticker symbols are valid
   - Check date range is reasonable

2. **Package Installation**
   - If encountering errors, install packages individually:
     ```bash
     pip install yfinance pandas numpy matplotlib openpyxl python-docx
     ```

3. **Output Generation**
   - Ensure write permissions in output directory
   - Close any open Excel/Word files before running
   - Verify sufficient disk space for figure generation

## Contributing
Feel free to submit issues and enhancement requests!

## License
This project is intended for educational purposes only. Use at your own risk.# Automated-Excel-Report-Generator
