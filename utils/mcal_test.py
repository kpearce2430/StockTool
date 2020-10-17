
import sys


import datetime
# from datetime import datetime, timedelta
# import pandas as pd
import numpy as np
import pandas_market_calendars as mcal

nyse = mcal.get_calendar('NYSE')
holidays = nyse.holidays()

# print(holidays.holidays)
nToday = np.datetime64('today')
myHolidays = []
for h in holidays.holidays:
    if nToday - h < np.timedelta64(365*3,'D') and nToday - h > np.timedelta64(0,'D'):
        print(h,nToday-h)
        myHolidays.append(h)

start_date = datetime.datetime(2019, 1, 1)
end_date = datetime.datetime(2019, 12, 31)
testDate = np.datetime64('2018-03-30')
testDate2 = np.datetime64(end_date)
if testDate in myHolidays:
    print("Yeah!")

if end_date in myHolidays:
    print("Boo!")
else:
    print("Double Yeah!")
#print(myHolidays)


# print(nyse.valid_days(start_date='2018-03-29',end_date='2020-09-30'))
# print(holidays)
