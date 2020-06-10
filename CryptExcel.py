from binance.client import Client
from datetime import datetime
from time import sleep
import openpyxl

#Read CryptExcel.config
f = open("CryptExcel.config", "r")
content = f.read().split("\n")
api_key = content[0].replace("api_key", "").replace("=", "").strip()
secret_key = content[1].replace("secret_key", "").replace("=", "").strip()
pairs = content[2].replace("pairs", "").replace("=", "").replace(" ", "").strip().split(",")
refresh = int(content[3].replace("refresh", "").replace("=", "").strip())

def modify():
    #Binance API
    client = Client(api_key, secret_key)
    balances = client.get_account().get('balances')
    prices = client.get_all_tickers()
    trades = []
    for pair in pairs:
        trade = (client.get_all_orders(symbol=pair, limit=1000))
        for order in trade:
            trades.append(order)

    #Excel Values
    workbook = openpyxl.load_workbook('trading.xlsx')
    Home = workbook.get_sheet_by_name('Home')
    DetailledHome = workbook.get_sheet_by_name('DetailledHome')
    Wallet = workbook.get_sheet_by_name('Wallet')
    values = workbook.get_sheet_by_name('Values')

    #Fill Wallet sheet
    for i in range(len(balances)):
        Wallet[f'A{i+2}'] = balances[i].get('asset')
        Wallet[f'B{i+2}'] = balances[i].get('free')
        Wallet[f'C{i+2}'] = balances[i].get('locked')
    Wallet['E1'] = "Last update at " + str(datetime.now().strftime("%H:%M:%S %d/%m/%Y"))


    #Fill DetailledHome sheet
    trade_id = []
    for i in range(2, 10000):
        val = DetailledHome[f'B{i}'].value
        if val != None:
            trade_id.append(val)

    for order in trades:
        if trade_id.count(order.get('orderId')) == 0:
            DetailledHome[f'A{len(trade_id)+2}'] = order.get('symbol')
            DetailledHome[f'B{len(trade_id)+2}'] = order.get('orderId')
            DetailledHome[f'C{len(trade_id)+2}'] = order.get('orderListId')
            DetailledHome[f'D{len(trade_id)+2}'] = order.get('clientOrderId')
            DetailledHome[f'E{len(trade_id)+2}'] = order.get('price')
            DetailledHome[f'F{len(trade_id)+2}'] = order.get('origQty')
            DetailledHome[f'G{len(trade_id)+2}'] = order.get('executedQty')
            DetailledHome[f'H{len(trade_id)+2}'] = order.get('cummulativeQuoteQty')
            DetailledHome[f'I{len(trade_id)+2}'] = order.get('status')
            DetailledHome[f'J{len(trade_id)+2}'] = order.get('timeInForce')
            DetailledHome[f'K{len(trade_id)+2}'] = order.get('type')
            DetailledHome[f'L{len(trade_id)+2}'] = order.get('side')
            DetailledHome[f'M{len(trade_id)+2}'] = order.get('stopPrice')
            DetailledHome[f'N{len(trade_id)+2}'] = order.get('icebergQty')
            DetailledHome[f'O{len(trade_id)+2}'] = datetime.fromtimestamp(int(str(order.get('time'))[:-3]))
            DetailledHome[f'P{len(trade_id)+2}'] = order.get('time')
            DetailledHome[f'Q{len(trade_id)+2}'] = datetime.fromtimestamp(int(str(order.get('updateTime'))[:-3]))
            DetailledHome[f'R{len(trade_id)+2}'] = order.get('updateTime')
            DetailledHome[f'S{len(trade_id)+2}'] = order.get('isWorking')
            DetailledHome[f'T{len(trade_id)+2}'] = order.get('origQuoteOrderQty')
            trade_id.append(order.get('orderId'))
        else:
            index = trade_id.index(order.get('orderId'))
            DetailledHome[f'A{index+2}'] = order.get('symbol')
            DetailledHome[f'B{index+2}'] = order.get('orderId')
            DetailledHome[f'C{index+2}'] = order.get('orderListId')
            DetailledHome[f'D{index+2}'] = order.get('clientOrderId')
            DetailledHome[f'E{index+2}'] = order.get('price')
            DetailledHome[f'F{index+2}'] = order.get('origQty')
            DetailledHome[f'G{index+2}'] = order.get('executedQty')
            DetailledHome[f'H{index+2}'] = order.get('cummulativeQuoteQty')
            DetailledHome[f'I{index+2}'] = order.get('status')
            DetailledHome[f'J{index+2}'] = order.get('timeInForce')
            DetailledHome[f'K{index+2}'] = order.get('type')
            DetailledHome[f'L{index+2}'] = order.get('side')
            DetailledHome[f'M{index+2}'] = order.get('stopPrice')
            DetailledHome[f'N{index+2}'] = order.get('icebergQty')
            DetailledHome[f'O{index+2}'] = datetime.fromtimestamp(int(str(order.get('time'))[:-3]))
            DetailledHome[f'P{index+2}'] = order.get('time')
            DetailledHome[f'Q{index+2}'] = datetime.fromtimestamp(int(str(order.get('updateTime'))[:-3]))
            DetailledHome[f'R{index+2}'] = order.get('updateTime')
            DetailledHome[f'S{index+2}'] = order.get('isWorking')
            DetailledHome[f'T{index+2}'] = order.get('origQuoteOrderQty')
    DetailledHome['V1'] = "Last update at " + str(datetime.now().strftime("%H:%M:%S %d-%m-%Y"))

    #Feel Home sheet
    for i in range(len(trade_id)):
        Home[f'A{i+2}'] = f'=DetailledHome!O{i+2}'
        Home[f'B{i+2}'] = f'=DetailledHome!A{i+2}'
        Home[f'C{i+2}'] = f'=DetailledHome!L{i+2}'
        Home[f'D{i+2}'] = round(float(DetailledHome[f'F{i+2}'].value)*float(DetailledHome[f'E{i+2}'].value), 4)
        Home[f'E{i+2}'] = f'=DetailledHome!F{i+2}'
        Home[f'G{i+2}'] = f'=DetailledHome!E{i+2}'
        Home[f'H{i+2}'] = f'=DetailledHome!I{i+2}'
    Home['K1'] = '=DetailledHome!V1'

    #Fill Values sheet
    for i in range(len(prices)):
        values[f'A{i+2}'] = prices[i].get('symbol')
        values[f'B{i+2}'] = prices[i].get('price')
    values['D1'] = "Last update at " + str(datetime.now().strftime("%H:%M:%S %d-%m-%Y"))

    #Save Excel
    try:
        workbook.save('trading.xlsx')
        print('Saved !')
    except:
        print('Not saved because file is open...')
        pass

while True:
    modify()
    sleep(refresh)