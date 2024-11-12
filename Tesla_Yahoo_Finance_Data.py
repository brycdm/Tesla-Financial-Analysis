import yfinance as yf
import pandas as pd

# Identifying ticker symbol
ticker_symbol = "TSLA"
tesla = yf.Ticker(ticker_symbol)

# Downloading financial data
income_statement = tesla.financials.T
balance_sheet = tesla.balance_sheet.T
cash_flow = tesla.cashflow.T

# Resetting index for each data frame to prepare for renaming
income_statement.reset_index(inplace=True)
balance_sheet.reset_index(inplace=True)
cash_flow.reset_index(inplace=True)

# Renaming columns for consistency
income_statement.rename(columns={"index": "Date"}, inplace=True)
balance_sheet.rename(columns={"index": "Date"}, inplace=True)
cash_flow.rename(columns={"index": "Date"}, inplace=True)

# Selecting only key columns for each statement based on available columns
try:
    # Income Statement Key Metrics
    income_statement_key_metrics = income_statement[['Date', 'Total Revenue', 'Cost Of Revenue', 'Gross Profit', 'Operating Income', 'Operating Expense']]

    # Balance Sheet Key Metrics
    balance_sheet_summary = balance_sheet[['Date', 'Total Assets', 'Total Liabilities Net Minority Interest', 'Stockholders Equity', 'Working Capital', 'Total Debt']]

    # Cash Flow Statement Key Metrics
    cash_flow_summary = cash_flow[['Date', 'Operating Cash Flow', 'Investing Cash Flow', 'Financing Cash Flow', 'Free Cash Flow', 'Capital Expenditure']]

    # Saving all filtered data to a single Excel workbook with separate sheets
    with pd.ExcelWriter("Tesla_Financial_Summary.xlsx") as writer:
        income_statement_key_metrics.to_excel(writer, sheet_name="Income Statement", index=False)
        balance_sheet_summary.to_excel(writer, sheet_name="Balance Sheet", index=False)
        cash_flow_summary.to_excel(writer, sheet_name="Cash Flow", index=False)

    print("Filtered data saved in Tesla_Financial_Summary.xlsx with separate sheets for each statement.")

except KeyError as e:
    print("One or more columns not found. Check the available columns and adjust if necessary.")

