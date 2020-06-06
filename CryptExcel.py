from binance.client import Client
from datetime import datetime
import openpyxl

#Binance API
api_key = "Your api key"
secret_key = "Your secret key"
client = Client(api_key, secret_key)
balances = client.get_account().get('balances')
prices = client.get_all_tickers()
pairs = ["ETHUSDT", "ETHEUR"] #Add in the array the other pairs for which you want to recover the trade
trades = []
for pair in pairs:
    trade = (client.get_all_orders(symbol=pair, limit=1000))
    for order in trade:
        trades.append(order)


#Excel Values
workbook = openpyxl.load_workbook('trading.xlsx')
DetailledHome = workbook.get_sheet_by_name('DetailledHome')
Wallet = workbook.get_sheet_by_name('Wallet')
values = workbook.get_sheet_by_name('Values')
alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']



while True:
    #Fill Wallet sheet
    y=2
    for balance in balances:
        x=0
        for i in balance:
            Wallet[alphabet[x]+str(y)] = balance.get(i)
            x += 1
        y += 1


    #Fill DetailledHome
    y=2
    for order in trades:
        x=0
        for i in order:
            if i == "time" or i == "updateTime":
                timestamp = int(str(order.get(i))[:-3])
                time = datetime.fromtimestamp(timestamp)
                DetailledHome[alphabet[x]+str(y)] = str(time)
                x+=1
            DetailledHome[alphabet[x]+str(y)] = order.get(i)
            x+=1
        y+=1

    #Fill Values
    for i in range(len(prices)):
        values[f'A{i+2}'] = prices[i].get('symbol')
        values[f'B{i+2}'] = prices[i].get('price')

    #Save Excel
    workbook.save('trading.xlsx')
