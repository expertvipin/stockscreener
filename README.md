# MoneyControl-Scraper [2022]
_Scrape Stocks/Funds data from moneycontrol into a .excel file without needing to register for any API access_

### How to use it?

Firstly, make sure you have selenium >= 3.141.0, GeckoDriver and FireFox installed.
Then,run these 3 commands :
pip3 install selenium
pip3 install pandas
pip3 install openpyxl



USE 'scraper.py' to collect your data and get saved in excel files.

usage : 

For Stocks :::
	python3 scraper.py [--stocks STOCKS [STOCKS ...]]

For Funds :::
	python3 scraper.py [--funds STOCKS [STOCKS ...]]


arguments:
  -h, --help            show this help message and exit
  -s, --stocks STOCKS [STOCKS ...]
                        List the stocks you want to scrape for details
  
  -f, --funds FUNDS [FUNDS ...]
                        List the mutual funds you want to scrape for details


 Example:  

	for stocks:::
	 ``` python3 scraper.py --stocks "Suzlon Energy" "State Bank of India" ```

	 for funds:::
	  ``` python3 scraper.py --funds "Nippon India Small Cap Fund" "Aditya Birla Sun Life Digital India Fund" ```

The output is `StocksData.xlsx` for stocks and `FundsData.xlsx` for funds inside the main folder.
