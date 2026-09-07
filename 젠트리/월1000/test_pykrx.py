# -*- coding: utf-8 -*-
from pykrx import stock
from datetime import datetime, timedelta

TICKERS = {
    "0094M0": "RISE 코리아밸류업위클리고정커버드콜",
    "0040Y0": "SOL 팔란티어커버드콜OTM채권혼합",
    "490600": "RISE 미국배당100데일리고정커버드콜",
    "083650": "비에이치아이",
}

today = datetime.now()
start = (today - timedelta(days=10)).strftime("%Y%m%d")
end = today.strftime("%Y%m%d")

for code, name in TICKERS.items():
    df = stock.get_market_ohlcv(start, end, code)
    print(f"\n=== {name} ({code}) ===")
    print(df.tail(3))
