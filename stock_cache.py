#
#
# import pprint
import json
import common_request
import time
import os
import sys
from base64 import b64encode
import datetime

# from datetime import date
import holidays
import numpy as np
import pandas_market_calendars as mcal


class StockCache:
    def __init__(self):
        self.nyseHolidays = []
        self.getHolidays()

    def getHolidays(self):

        if len(self.nyseHolidays) == 0:
            nyse = mcal.get_calendar("NYSE")
            holidays = nyse.holidays()

            # print(holidays.holidays)
            nToday = np.datetime64("today")
            for h in holidays.holidays:
                # holidays has a larger number of holidays in a list.  We're only interested
                # in the last 3 years
                if nToday - h < np.timedelta64(
                    365 * 3, "D"
                ) and nToday - h > np.timedelta64(0, "D"):
                    # print(h, nToday - h)
                    self.nyseHolidays.append(h)

    def isHoliday(self, d):
        if len(self.nyseHolidays) == 0:
            self.getHolidays()

        if d in self.nyseHolidays:
            return True
        else:
            return False

    def isLeapYear(self, year=0):
        if (year % 4) == 0:
            if (year % 100) == 0:
                if (year % 400) == 0:
                    # print("{0} is a leap year".format(year))
                    return True
                else:
                    # print("{0} is not a leap year".format(year))
                    return False
            else:
                # print("{0} is a leap year".format(year))
                return True

        # print("{0} is not a leap year".format(year))
        return False

    def adjustForHolidaysAndWeekends(self, jDate):

        if len(self.nyseHolidays) == 0:
            self.getHolidays()

        dtDate = datetime.datetime.strptime(jDate, "%Y%j").date()
        dtTime = dtDate.timetuple()
        myYear = int(jDate[0:4])
        myJul = int(jDate[4:7])
        # print(">> {}:{}".format(myYear, myJul))

        while True:
            # Saturday or Sunday
            if dtTime.tm_wday <= 4:
                # print( "weekday [{}] - {}:{}:{}".format( dtTime.tm_wday, dtTime.tm_year, dtTime.tm_mon, dtTime.tm_mday ) )
                if ( datetime.date(dtTime.tm_year, dtTime.tm_mon, dtTime.tm_mday) not in self.nyseHolidays ):
                    # print("not holiday {}".format(jDate))
                    return jDate
                # else:
                #    print("is holiday {}".format(jDate))
                # else:
                # print( "weekend [{}] - {}:{}:{}".format( dtTime.tm_wday, dtTime.tm_year, dtTime.tm_mon, dtTime.tm_mday ) )

            myJul = myJul - 1
            if myJul < 1:
                myYear = myYear - 1
                if self.isLeapYear(myYear):
                    myJul = myJul + 366
                else:
                    myJul = myJul + 365

            # there may be a better way to do this
            jDate = str(myYear) + str(myJul).zfill(3)
            dtDate = datetime.datetime.strptime(jDate, "%Y%j").date()
            dtTime = dtDate.timetuple()
            myYear = int(jDate[0:4])
            myJul = int(jDate[4:7])

    def couchURL(self):
        if hasattr(self, "couchdb_url") == False:
            self.couchdb_url = os.getenv("COUCHDB_URL", "http://localhost:5984")

        return self.couchdb_url

    def couchHeader(self):
        if hasattr(self, "headers") == False:
            # TODO Store Admin credentials securely
            user = "admin"
            password = "password"
            self.headers = {
                "Authorization": "Basic {}".format(
                    b64encode(bytes(f"{user}:{password}", "utf-8")).decode("ascii")
                ),
                "Content-type": "application/json",
            }
        # print(headers)
        return self.headers

        # Does the work for reading data out of a database

    def couchDataRead(self, ticker, jDate, database):
        # in quote?
        url = self.couchURL() + "/" + database + "/" + ticker.upper() + ":" + jDate
        conn = common_request.CommonRequest("GET", url, self.couchHeader())

        #
        if conn.Request():
            # print("Success:", conn)
            #  Yes There is something in the db,
            #  Convert it to JSON and return
            data = json.loads(conn.ResponseData())
            conn.Close()
            return data

        # print("Unable to Read Data", conn)
        return None

    # Does the work to save data to a database.
    def couchDataSave(self, ticker, jDate, dataBase, dataRecord, rev=None):
        dataRecord["_id"] = ticker.upper() + ":" + jDate
        dataRecord["key"] = ticker.upper() + ":" + jDate
        if rev != None:
            dataRecord["_rev"] = rev
        dataRecord["symbol"] = ticker.upper()
        dataRecord["julian"] = jDate

        # print("Saving {}:{}:{}".format(ticker.upper(), jDate, dataRecord))

        url = self.couchURL() + "/" + dataBase + "/" + ticker.upper() + ":" + jDate

        conn = common_request.CommonRequest("PUT", url, self.couchHeader(), dataRecord)
        if conn.Request("PUT", url, self.couchHeader(), dataRecord):
            # print("Successfully inserted", conn)
            return True

        print("Failed:", conn)
        return False

    def iexStockHistoryGet(self, ticker, types="chart&range=5y", token=None):
        # def GetIEXStockHistory(ticker, types="chart&range=5y", token=None):

        if len(ticker) > 4:
            # print("Invalid Ticker [{}] for IEX".format(ticker))
            return

        if hasattr(self, "token") == False:
            self.token = os.getenv("TOKEN", "junk")

        #  example'https://cloud.iexapis.com/v1/stock/HD/batch?types=chart&range=5y&token=<blah>
        url = (
            "https://cloud.iexapis.com/v1/stock/"
            + ticker.lower()
            + "/batch?types="
            + types
            + "&token="
            + self.token
        )

        conn = common_request.CommonRequest("GET", url)

        if conn.Request():
            # print("Success:", conn)
            data = json.loads(conn.ResponseData())
            conn.Close()
            return data

        print("ERROR Failure:", conn)
        conn.Close()
        return None

    # def SaveIEXHistoryData(ticker, histData):
    def iexStockHistorySave(self, ticker, histData):

        if not isinstance(histData, dict):
            print("ERROR: Historical Data is not json")
            return None

        print("iexHistoryDataSave: {}".format(ticker))
        charts = histData.get("chart")

        #
        for c in charts:
            #
            chartDate = c.get("date")

            # Get the Julian Date and Year
            chartDT = datetime.datetime.strptime(chartDate, "%Y-%m-%d")
            chartTT = chartDT.timetuple()
            jul = str(chartTT.tm_yday)
            jDate = str(chartTT.tm_year) + jul.zfill(3)
            # print("{}:{}:{}".format(ticker, chartDate, jDate))

            # See if there is a record
            response = self.historyDataRead(ticker, jDate)
            if isinstance(response, dict):
                iexHistory = response.get("iex_history")
                if iexHistory != None:
                    # print("History already exists")
                    continue  # were not going to overwrite it
                else:
                    revision = response.get("_rev")
            else:
                revision = None
                response = {}

            # at this point we either have a new record, or
            # a record without a 'iex_history' record.
            response["iex_history"] = c
            print("Saving History rev[{}]",revision)
            self.historyDataSave(ticker, jDate, response, revision)

    def iexStockQuoteGet(
        self, ticker, types="quote,stats,news,dividends&range=1y&last=3", token=None
    ):

        if len(ticker) > 4:
            print("Invalid Ticker [{}] for IEX".format(ticker))

        if hasattr(self, "token") == False:
            self.token = os.getenv("TOKEN", "junk")

        # "/batch?types=quote,stats,news,dividends&range=1y&last=3&token="
        url = (
            "https://cloud.iexapis.com/v1/stock/"
            + ticker.lower()
            + "/batch?types="
            + types
            + "&token="
            + self.token
        )

        conn = common_request.CommonRequest("GET", url)

        if conn.Request():
            # print("Success:", conn)
            data = json.loads(conn.ResponseData())
            conn.Close()
            return data

        print("Failure:", conn)
        conn.Close()
        return None

    # wrapper function to read quote database
    def quoteDataRead(self, ticker, jDate):
        return self.couchDataRead(ticker, jDate, "quotes")

    # wrapper function for the quotes database
    def quoteDataSave(self, ticker, jDate, quoteData, rev=None):
        return self.couchDataSave(ticker, jDate, "quotes", quoteData, rev)

    # wrapper function for the history database
    def historyDataSave(self, ticker, jDate, historyData, rev=None):
        return self.couchDataSave(ticker, jDate, "historical", historyData, rev)

    def historyDataRead(self, ticker, jDate):

        # print("JDate In:{}", jDate)
        jDate = self.adjustForHolidaysAndWeekends(jDate)
        # print("JDate Out:{}", jDate)
        return self.couchDataRead(ticker, jDate, "historical")



    def jDateFromTime(self, tt):
        if isinstance(tt, time.struct_time):
            jul = str(tt.tm_yday)
            return str(tt.tm_year) + jul.zfill(3)

        return None

    def timeFromJul(self, jul):
        return time.strptime(jul, "%Y%j")

    #
    # * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    #
    def StockLookup(self, ticker, **kwargs):

        if len(ticker) > 4:
            print("{} is not a stock".format(ticker))
            return None

        tt = time.localtime()
        jDateToday = self.jDateFromTime(tt)

        jDate = kwargs.get("jul_date")
        tDate = kwargs.get("time")

        if jDate == None and tDate == None:
            # Nothing of interest provided
            # tDate = tt
            jDate = jDateToday

        else:
            # Something was passed in
            if tDate != None and isinstance(tDate, time.struct_time):
                # print("tDate:{}".format(str(tDate)))
                jDate = self.jDateFromTime(tDate)
            # elif jDate != None:
            #    tDate = time.strftime(jDate,"%Y%j")

        # print("jDate:{}".format(jDate))
        # response = None  # by default
        if jDate == None or jDate == jDateToday:
            # Get todays quote
            if jDate == None:
                jDate = jDateToday

            response = self.quoteDataRead(ticker, jDate)
            revision = None

            # is quote current?
            if response != None:
                # print("Success:", conn)
                #  Yes There is something in the db
                quoteData = response.get("quote")
                if isinstance(quoteData, dict):
                    # Check the market open flag, If it's true, the market was
                    # open (9:30a-4:00p) when the request was made.
                    # TODO:  Make a request prior to market open to capture the result.  Not sure what the
                    #  'is US Market Open' flag is set to during this time.
                    # May need to modify code to look if it's before 9:30, then get the prior day's quote.
                    #
                    marketOpen = quoteData.get("isUSMarketOpen")
                    closePrice = quoteData.get("close")
                    # print("isUSMarketOpen: {}".format(str(marketOpen)))
                    if closePrice != 'None' and marketOpen == False:  # should be False
                        # print("Market not open, return saved result")
                        # at this point, the data in 'quotes' was current.
                        # Just return it to the user
                        return response

                    # Get the revision number, we're going to make
                    # another request and update the record
                    revision = response.get("_rev")

                else:
                    print("WARNING! Unexpected result[{}]".format(quoteData))

            print(revision)

            # Go to IEX and get a quote.
            iexData = self.iexStockQuoteGet(ticker)
            if iexData != None:
                #
                # Let's save it off

                if self.quoteDataSave(ticker, jDate, iexData, revision) == False:
                    print("ERROR Saving Quote Data")

                # We're only saving off the quote data in history
                quoteData = iexData.get("quote")

                if isinstance(quoteData, dict):
                    # quoteData["Source"] = "IEX"

                    histData = self.historyDataRead(ticker, jDate)
                    if isinstance(histData, dict):
                        revision = histData.get("_rev")
                    else:
                        histData = {}
                        revision = None

                    histData["iex_quote"] = quoteData
                    if self.historyDataSave(ticker, jDate, histData, revision) == False:
                        print("Error Saving History Data {}", histData)

                    return histData

                print("ERROR: Invalid response {}".format(iexData))

            else:
                print("ERROR: Call to IEX Failed")

        else:  # Not today,

            print("Here - Not Today")
            #  Let' see if it's in the quote data
            response = self.historyDataRead(ticker, jDate)

            print("Here>>>:", response)

            if isinstance(response,dict):
                quoteData = response.get("iex_quote")
                if isinstance(quoteData, dict):
                    # Check the market open flag, If it's true, the market was
                    # open (9:30a-4:00p) when the request was made.
                    # TODO:  Make a request prior to market open to capture the result.  Not sure what the
                    #  'is US Market Open' flag is set to during this time.
                    # May need to modify code to look if it's before 9:30, then get the prior day's quote.
                    #
                    marketOpen = quoteData.get("isUSMarketOpen")
                    closePrice = quoteData.get("close")
                    print("marketOpen: {} close: {}".format(marketOpen,closePrice))

                    if marketOpen == False and closePrice != None:
                        print("Returng reponse based in iex_quote")
                        return response

                historyData = response.get("iex_history")
                if isinstance(historyData, dict):
                    # TODO: Make sure this has valid data
                    print("Returng reponse based in iex_history")
                    return response

            # Nothing found in history database, let's check the history.
            print("Checking Quote Data")
            response = self.quoteDataRead(ticker, jDate)

            if isinstance(response,dict):
                quoteData = response.get("quote")
                if isinstance(quoteData, dict):
                    # Check the market open flag, If it's true, the market was
                    # open (9:30a-4:00p) when the request was made.
                    # TODO:  Make a request prior to market open to capture the result.  Not sure what the
                    #  'is US Market Open' flag is set to during this time.
                    # May need to modify code to look if it's before 9:30, then get the prior day's quote.
                    #
                    marketOpen = quoteData.get("isUSMarketOpen")
                    closePrice = quoteData.get("close")

                    if marketOpen == False and closePrice != None:
                        return response

            # Nothing or Invalid found in both databases.
            # Print in most cases if this is a Saturday or Sunday
            #
            print("No history or quote found for {} on {}".format(ticker, jDate))
            #
            response = self.iexStockHistoryGet(ticker)
            if isinstance(response, dict):
                # print("Saving: {}", response)
                self.iexStockHistorySave(ticker, response)

                return self.historyDataRead(ticker, jDate)

        # Note, that if a record was found in quote and
        # we failed to get an updated version, the original will
        # be returned
        print("Returning {}".format(response))
        return response


#
#  This is the original function
#
def CacheLookup(ticker, jDate=None):

    cache = StockCache()

    if jDate == None:
        tt = time.localtime()
        jul = str(tt.tm_yday)
        jDate = str(tt.tm_year) + jul.zfill(3)

    url = cache.couchURL() + "/quotes/" + ticker.upper() + ":" + jDate

    # hdr = myHeader()
    # print(hdr)

    conn = common_request.CommonRequest("GET", url, cache.couchHeader())

    if conn.Request():
        # print("Success:", conn)
        return json.loads(conn.ResponseData())

    print("WARNING: Failed Cache Lookup:", conn)

    myData = cache.iexStockQuoteGet(ticker)
    if isinstance(myData, dict):
        cache.quoteDataSave(ticker, jDate, myData)
        return myData
    else:
        print("ERROR: Failed GetIEXStockQuote: ", str(myData))

    return None


"""

def testDates():
    years = [2019, 2020, 2021]
    for y in years:
        if isLeapYear(y):
            numDays = 367
        else:
            numDays = 366

        for i in range(1, numDays):
            jDate = str(y).zfill(4) + str(i).zfill(3)
            print("{}:{}".format(jDate, adjustForHolidaysAndWeekends(jDate)))
"""

# main guy if running
if __name__ == "__main__":

    cache = StockCache()

    stocks = ["JNJ", "AAPL", "hd"]
    for s in stocks:
        print("No Arg {}:{}".format(s, cache.StockLookup(s)))
    #
    # stocks.append("JUNK")

    stocks = ["CSX", "AAPL"]
    # jDate = "2018090"  # Known Saturday
    mDate = datetime.date(2020,9,30)
    for s in stocks:
        # print("jul_date {}:{}".format(s, cache.StockLookup(s, jul_date=jDate)))
        print("mDate {}:{}".format(mDate, cache.StockLookup(s, time=mDate.timetuple())))


    # for s in stocks:
    #    myTime =  timeFromJul("2020123")
    #    print("time {}:{}".format(s,StockLookup(s,time=myTime)))
