import yfinance as yf
import os
from odf.opendocument import OpenDocumentSpreadsheet, load
from odf.table import Table, TableRow, TableCell
from odf.text import P

# Function to fetch stock prices from Yahoo Finance
def fetch_stock_prices(company_name, nse_symbol, bse_symbol):
    try:
        # Fetch NSE price
        nse_stock = yf.Ticker(nse_symbol)
        nse_price = nse_stock.history(period="1d")['Close'].iloc[-1]

        # Fetch BSE price
        bse_stock = yf.Ticker(bse_symbol)
        bse_price = bse_stock.history(period="1d")['Close'].iloc[-1]

        return nse_price, bse_price
    except Exception as e:
        print(f"Error fetching data: {e}")
        return None, None

# Function to save data into ODS
def save_to_ods(company_name, nse_price, bse_price, filename="stock_prices.ods"):
    if os.path.exists(filename):
        # Load existing ODS file
        ods_doc = load(filename)
        spreadsheet = ods_doc.spreadsheet
    else:
        # Create a new ODS file if it doesn't exist
        ods_doc = OpenDocumentSpreadsheet()
        spreadsheet = ods_doc.spreadsheet

    # Check if table already exists; if not, create one
    tables = spreadsheet.getElementsByType(Table)
    if len(tables) == 0:
        table = Table(name="Stock Prices")
        spreadsheet.addElement(table)
    else:
        table = tables[0]

    # Add the row with company name, NSE price, and BSE price
    row = TableRow()
    for value in [company_name, nse_price, bse_price]:
        cell = TableCell()
        cell.addElement(P(text=str(value)))
        row.addElement(cell)
    table.addElement(row)
    
    # Save the file
    ods_doc.save(filename)
    print(f"Data saved to {filename}")

if __name__ == "__main__":

    # Input: Company Name and its corresponding NSE and BSE symbols
    company_name = input("Enter the company name: ")
    nse_symbol = input("Enter the NSE symbol : ")
    bse_symbol = input("Enter the BSE symbol : ")

    # Fetch the stock prices
    nse_price, bse_price = fetch_stock_prices(company_name, nse_symbol, bse_symbol)

    # If prices were fetched successfully, save them into an ODS
    if nse_price is not None and bse_price is not None:
        save_to_ods(company_name, nse_price, bse_price)