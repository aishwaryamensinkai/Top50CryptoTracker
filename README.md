# Cryptocurrency Analysis Project

## Overview
This project is designed to fetch live cryptocurrency data for the top 50 cryptocurrencies, analyze it, and present the data in a live-updating Excel sheet. Additionally, an analysis report is generated summarizing key insights from the fetched data.

## Features
- Fetches live data from the CoinMarketCap API for the top 50 cryptocurrencies.
- Provides key metrics such as current price, market capitalization, 24-hour trading volume, and 24-hour price change.
- Analyzes and identifies the top 5 cryptocurrencies by market cap, the average price, and the highest/lowest 24-hour price changes.
- Updates the data in an Excel sheet every 5 minutes.
- Generates an analysis report in PDF format.

## Prerequisites
Ensure you have Python 3.6+ installed on your machine. You will also need the following libraries:

- `requests`: For making API requests
- `pandas`: For data handling and analysis
- `openpyxl`: For updating the Excel sheet
- `fpdf`: For generating the PDF report

You can install the necessary libraries using the following command:

```bash
pip install requests pandas openpyxl fpdf
```

## Getting Started
1. Clone the repository or download the script files.
2. Obtain your API key from the [CoinMarketCap API](https://coinmarketcap.com/api/).

### Setting up your API Key
- Replace `'your_api_key'` in the script with your actual API key:

```python
api_key = 'your_api_key'
```

## Running the Script
1. Open your terminal or command prompt.
2. Navigate to the directory where the script is located.
3. Run the script using the following command:

```bash
python crypto_data_live.py
```

The script will start fetching live cryptocurrency data and will update the Excel file (`Live_Crypto_Data.xlsx`) every 5 minutes.

## Viewing the Analysis Report
The analysis report is generated in real-time and saved as `Cryptocurrency_Analysis_Report.pdf` in the project directory. It summarizes the key insights such as top 5 cryptocurrencies by market cap, average price, and the highest/lowest 24-hour price changes.

## Output Files
- `Live_Crypto_Data.xlsx`: Contains the continuously updated cryptocurrency data.
- `Enhanced_Cryptocurrency_Analysis_Report.pdf`: A detailed PDF report generated from the live data.

## Troubleshooting
- If the data doesn't fetch correctly, check your internet connection and ensure your API key is valid.
- Ensure all required Python libraries are installed.

## License
This project is licensed under the MIT License. Feel free to modify and distribute as needed.

## Contact
For any queries or issues, please contact Aishwarya Girish Mensinkai at aishwaryamensinkai@gmail.com.

