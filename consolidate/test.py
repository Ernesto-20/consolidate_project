import locale
#
print(locale.getlocale())
locale.setlocale(locale.LC_ALL, 'es_CU')
print(locale.getlocale())
print(locale.currency(123131231.311, grouping=True, symbol=True))

import datetime
Current_Date = datetime.datetime.now()
# print (str(Current_Date.hour) + ' and ' + str(Current_Date.minute))

import pandas as pd
# print('VACIO: ', pd.NA, pd.NaT)


