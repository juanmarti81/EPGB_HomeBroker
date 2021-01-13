import pandas as pd
from datetime import date, timedelta
import xlwings as xw

wb = xw.Book('EPGB V3.3.1 FE.xlsb')
shtTickers = wb.sheets('Tickers')
bases_GGAL = pd.DataFrame()
bases_PAM = pd.DataFrame()
bases_COME = pd.DataFrame()
allOptions = pd.DataFrame()
oOpciones = pd.DataFrame()

def getOptionsList():
    global allOptions
    rng = shtTickers.range('A2:A500')
    oOpciones = rng.value
    allOptions = pd.DataFrame({'symbol': oOpciones},
                              columns=["symbol", "bid_size", "bid", "ask", "ask_size", "last",
                                       "change", "open", "high", "low", "previous_close", "turnover", "volume",
                                       'operations', 'datetime'])
    allOptions = allOptions.set_index('symbol')
    allOptions['datetime'] = pd.to_datetime(allOptions['datetime'])

    return allOptions

def getAccionesList():
    rng = shtTickers.range('D2:D500').expand()
    oAcciones = rng.value
    ACC = pd.DataFrame({'symbol' : oAcciones}, columns=["symbol", "bid_size", "bid", "ask", "ask_size", "last",
                                                                             "change", "open", "high", "low", "previous_close", "turnover", "volume",
                                                                             'operations', 'datetime'])
    ACC = ACC.set_index('symbol')
    ACC['datetime'] = pd.to_datetime(ACC['datetime'])
    return ACC

def getBonosList():
    rng = shtTickers.range('F2:F500').expand()
    oBonos = rng.value
    Bonos = pd.DataFrame({'symbol' : oBonos}, columns=["symbol", "bid_size", "bid", "ask", "ask_size", "last",
                                                                             "change", "open", "high", "low", "previous_close", "turnover", "volume",
                                                                             'operations', 'datetime'])
    Bonos = Bonos.set_index('symbol')
    Bonos['datetime'] = pd.to_datetime(Bonos['datetime'])
    return Bonos

def getCedearsList():
    rng = shtTickers.range('H2:H500').expand()
    oCedears = rng.value
    Cedears = pd.DataFrame({'symbol' : oCedears}, columns=["symbol", "bid_size", "bid", "ask", "ask_size", "last",
                                                                             "change", "open", "high", "low", "previous_close", "turnover", "volume",
                                                                             'operations', 'datetime'])
    Cedears = Cedears.set_index('symbol')
    Cedears['datetime'] = pd.to_datetime(Cedears['datetime'])
    return Cedears

# Cauciones
i = 1
fechas = []
while i < 31:
    fecha = date.today() + timedelta(days=i)
    fechas.extend([fecha])
    i += 1

cauciones = pd.DataFrame({'settlement':fechas}, columns=['settlement','last', 'turnover', 'bid_amount', 'bid_rate', 'ask_rate', 'ask_amount'])
cauciones['settlement'] = pd.to_datetime(cauciones['settlement'])
cauciones = cauciones.set_index('settlement')


