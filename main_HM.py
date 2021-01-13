from pyhomebroker import HomeBroker
import xlwings as xw
import Options_Helper_HM
import pandas as pd
import time
import config

ACC = Options_Helper_HM.getAccionesList()
cedears = Options_Helper_HM.getCedearsList()
cauciones = Options_Helper_HM.cauciones
options = Options_Helper_HM.getOptionsList()
bonos = Options_Helper_HM.getBonosList()
everything = ACC.append(bonos)
everything = everything.append(cedears)
listLength = len(everything) + 2

# Hojas del excel
wb = xw.Book('EPGB V3.3.1 FE.xlsb')
shtTest = wb.sheets('HomeBroker')
shtTickers = wb.sheets('Tickers')


def on_options(online, quotes):
    global options
    thisData = quotes
    thisData = thisData.drop(['expiration', 'strike', 'kind'], axis=1)
    thisData['change'] = thisData["change"] / 100
    thisData['datetime'] = pd.to_datetime(thisData['datetime'])
    options.update(thisData)


def on_securities(online, quotes):
    global ACC
    thisData = quotes
    thisData = thisData.reset_index()
    thisData['symbol'] = thisData['symbol'] + ' - ' + thisData['settlement']
    thisData = thisData.drop(["settlement"], axis=1)
    thisData = thisData.set_index("symbol")
    thisData['change'] = thisData["change"] / 100
    thisData['datetime'] = pd.to_datetime(thisData['datetime'])
    everything.update(thisData)


def on_repos(online, quotes):
    global cauciones
    thisData = quotes
    thisData = thisData.reset_index()
    thisData = thisData.set_index("symbol")
    thisData = thisData[['PESOS' in s for s in quotes.index]]
    thisData = thisData.reset_index()
    thisData['settlement'] = pd.to_datetime(thisData['settlement'])
    thisData = thisData.set_index("settlement")
    thisData['last'] = thisData["last"] / 100
    thisData['bid_rate'] = thisData["bid_rate"] / 100
    thisData['ask_rate'] = thisData["ask_rate"] / 100
    thisData = thisData.drop(
        ['open', 'high', 'low', 'volume', 'operations', 'datetime'], axis=1)
    thisData = thisData[[
        'last', 'turnover', 'bid_amount', 'bid_rate', 'ask_rate', 'ask_amount'
    ]]
    cauciones.update(thisData)


def on_error(online, error):
    print("Error Message Received: {0}".format(error))


hb = HomeBroker(int(config.broker),
                on_options=on_options,
                on_securities=on_securities,
                on_repos=on_repos,
                on_error=on_error)

hb.auth.login(dni=config.dni,
              user=config.user,
              password=config.password,
              raise_exception=True)
hb.online.connect()

# Comentar las líneas a lo que no se quiera recibir datos
hb.online.subscribe_options()
hb.online.subscribe_securities('bluechips', '48hs')
hb.online.subscribe_securities('bluechips', 'SPOT')
hb.online.subscribe_securities('government_bonds', '48hs')
hb.online.subscribe_securities('government_bonds', 'SPOT')
hb.online.subscribe_securities('cedears', '48hs')
hb.online.subscribe_securities('cedears', 'SPOT')
hb.online.subscribe_repos()

everything_last = pd.DataFrame()
options_last = None
cauciones_last = None

while True:
    try:
        oRange = 'A' + str(listLength)
        shtTest.range('A1').options(index=True, header=True).value = everything
        shtTest.range(oRange).options(index=True, header=False).value = options
        shtTest.range('S2').options(index=True, header=False).value = cauciones
        time.sleep(2)

    except:
        print('Hubo un error al actualizar excel')