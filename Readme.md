# Conexión con HomeBroker usando Python y planilla EPGB

## Instalación: ##

1. Instalar depedencias utilizando "pip install -r requirements.txt"
2. Crear archivo config.py
   Escribir las credenciales:
    ```
    broker = 00
    dni = '00000000'
    user = 'usuario'
    password = 'ásswprd'
    ```
    La lista de brokers y su código puede obtenerse en https://github.com/crapher/pyhomebroker

## Configuraciones: ##

Para reducir la cantidad de información que se envía a Excel, los tickers que se quieren deben listarse en la hoja "Tickers". Se debe escribir al menos 2 tickers para cada tipo.

Si hay información que no van a utilizar, lo ideal es comentar las líneas que correspondan:

```
hb.online.subscribe_options()
hb.online.subscribe_securities('bluechips', '48hs')
hb.online.subscribe_securities('bluechips', 'SPOT')
hb.online.subscribe_securities('government_bonds', '48hs')
hb.online.subscribe_securities('government_bonds', 'SPOT')
hb.online.subscribe_securities('cedears', '48hs')
hb.online.subscribe_securities('cedears', 'SPOT')
hb.online.subscribe_repos()
````

Los datos se actualizan cada 2 segundos. Puede cambiarse desde la siguiente línea:
```
 time.sleep(2)
```

pyHomeBroker fue creado por Diego Degese: https://github.com/crapher/pyhomebroker

Planilla creada por Guillermo Cutella @gcutte