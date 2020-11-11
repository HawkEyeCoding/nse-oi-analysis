#!/usr/bin/env python3

import requests
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta, TH
import google_sheet as gs
import dataframe_image as dfi


# Find next thursday to figure out weekly expiry.
def next_thu_expiry_date():
    todayte = datetime.today()

    next_thursday_expiry = todayte + relativedelta(weekday=TH(1))

    str_next_thursday_expiry = str(next_thursday_expiry.strftime("%d")) + "-" + next_thursday_expiry.strftime(
        "%b") + "-" + next_thursday_expiry.strftime("%Y")
    print(str_next_thursday_expiry)
    return str_next_thursday_expiry


def find_ce_pe(spot_price, ce_values, pe_values):
    spot_price = round(spot_price, -2)
    otm_calls = []
    otm_puts = []

    ce_dt = pd.DataFrame(ce_values).sort_values(['strikePrice'])
    pe_dt = pd.DataFrame(pe_values).sort_values(['strikePrice'], ascending=False)

    for ce_value in ce_dt['strikePrice']:
        if ce_value >= spot_price:
            otm_calls.append(ce_value)
            if len(otm_calls) > 10:
                break

    for ce_value in pe_dt['strikePrice']:
        if ce_value <= spot_price:
            otm_puts.append(ce_value)
            if len(otm_puts) > 10:
                break

    return otm_calls, otm_puts


def read_oi(spot_price, ce_values, pe_values, symbol):
    sheet_name = str(datetime.today().strftime("%d-%m-%Y")) + symbol

    strike_oi_ce_pe = find_ce_pe(float(spot_price), ce_values, pe_values)
    values = gs.read_row_values('master', 'Date', symbol, [str(datetime.today().strftime("%d-%m-%Y")), spot_price,
                                                           str(strike_oi_ce_pe[0][10]),
                                                           str(strike_oi_ce_pe[1][10])])
    strike_oi_ce_pe = list(strike_oi_ce_pe)
    if values == -1:
        gs.create_new_worksheet(sheet_name)
    else:
        if symbol == 'BANKNIFTY':
            spot_price_org = values[4]
        else:
            spot_price_org = values[1]
        strike_oi_ce_pe = find_ce_pe(float(spot_price_org), ce_values, pe_values)

    ce_dt = pd.DataFrame(ce_values).sort_values(['strikePrice'])
    pe_dt = pd.DataFrame(pe_values).sort_values(['strikePrice'], ascending=False)

    drop_col = ['pchangeinOpenInterest', 'askQty', 'askPrice', 'underlyingValue', 'totalBuyQuantity',
                'totalSellQuantity', 'bidQty', 'bidprice', 'change', 'pChange', 'impliedVolatility',
                'totalTradedVolume']
    ce_dt.drop(drop_col, inplace=True, axis=1)
    pe_dt.drop(drop_col, inplace=True, axis=1)

    ce_dt = ce_dt.loc[ce_dt['strikePrice'].isin(strike_oi_ce_pe[0])]
    pe_dt = pe_dt.loc[pe_dt['strikePrice'].isin(strike_oi_ce_pe[1])]

    # print(ce_dt[['strikePrice', 'lastPrice']])
    # print(pe_dt[['strikePrice', 'changeinOpenInterest']])

    print(ce_dt.head(1))
    print(pe_dt.head(1))

    gs.insert_record(sheet_name, spot_price, (ce_dt['changeinOpenInterest'].sum() * 75),
                     (pe_dt['changeinOpenInterest'].sum() * 75))

    # Read sheet and store and image to be sent to telegram
    dataframe = gs.read_sheet_into_df(sheet_name)
    dfi.export(dataframe, '/tmp/df_styled.png')
    file = open('/tmp/df_styled.png', 'rb')

    print(str([spot_price, (ce_dt['changeinOpenInterest'].sum() * 75),
               (pe_dt['changeinOpenInterest'].sum() * 75)]))


def main():
    symbols = ["BANKNIFTY", "NIFTY"]
    base_url = 'https://www.nseindia.com/api/option-chain-indices?symbol='
    url_oc = "https://www.nseindia.com/option-chain"

    for symbol in symbols:
        headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, ''like Gecko) '
                                 'Chrome/80.0.3987.149 Safari/537.36',
                   'accept-language': 'en,gu;q=0.9,hi;q=0.8', 'accept-encoding': 'gzip, deflate, br'}
        session = requests.Session()
        request = session.get(url_oc, headers=headers, timeout=5)
        cookies = dict(request.cookies)
        response = session.get(base_url + symbol, headers=headers, timeout=5, cookies=cookies)
        print(response.status_code)
        dajs = response.json()
        expiry = next_thu_expiry_date()

        spot_price = dajs['records']['underlyingValue']
        ce_values = [data['CE'] for data in dajs['records']['data'] if "CE" in data and data['expiryDate']
                     == expiry]
        pe_values = [data['PE'] for data in dajs['records']['data'] if "PE" in data and data['expiryDate']
                     == expiry]

        read_oi(spot_price, ce_values, pe_values, symbol)

        print(len(dajs))
        print("==================")


if __name__ == '__main__':
    main()
