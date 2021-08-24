import time

import requests

marketDataUrl = "https://query1.finance.yahoo.com/v7/finance/quote?symbols="
headers = {'User-Agent': 'X'}


class MarketData:
    def __init__(self, symbol, bid, ask):
        self.symbol, self.bid, self.ask = symbol, bid, ask

    def __repr__(self):
        return "[symbol=" + self.symbol + ", bid=" + str(self.bid) + ", ask=" + str(self.ask) + "]"


def get_market_data(stocks):
    stocks_delimited_by_comma = ",".join(stocks)
    final_url = marketDataUrl + stocks_delimited_by_comma
    response = requests.get(final_url, headers=headers)
    result_list = response.json()["quoteResponse"]["result"]
    market_data_result = []
    for element in result_list:
        market_data_result.append(MarketData(element["symbol"], element["bid"], element["ask"]))
    return market_data_result


if __name__ == '__main__':
    while 1 == 1:
        market_data = get_market_data(["AAPL", "TSLA", "F"])
        print(market_data)
        time.sleep(1)
