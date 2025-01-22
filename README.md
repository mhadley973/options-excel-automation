This program is meant as an organizational tool to aid people who trade options using Charles Schwab / E*TRADE
It uses the E*TRADE and Charles Schwab API to automatically populate a spreadsheet with your positions
The spreadsheet calculates your risks, potential gains and losses

You will need to create an API account with E*TRADE and Charles Schwab to use this program. You will need to update the .env file with your API keys.

You will also need to install the requirements.txt file using pip install -r requirements.txt in the terminal.

You also need to provide the account hash number for your desired TDA account on line 66 of tda_api.py

You also need to provide the account hash number for your desired E*TRADE accounts on line 24 of etrade_api.py
