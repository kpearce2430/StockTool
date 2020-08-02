#
#
import sys
import json
import common_request
import time
from base64 import b64encode


def myHeader():
    user = "admin"
    password = "password"
    headers = {
        "Authorization": "Basic {}".format(
            b64encode(bytes(f"{user}:{password}", "utf-8")).decode("ascii")
        ),
        "Content-type": "application/json",
    }
    # print(headers)
    return headers


def GetStockQuote(ticker):

    token = "pk_189dd9a1c5814706a37220a212dc54a0"
    # "/batch?types=quote,stats,news,dividends&range=1y&last=3&token="
    url = (
        "https://cloud.iexapis.com/v1/stock/"
        + ticker.lower()
        + "/batch?types=quote,stats,news,dividends&range=1y&last=3&token="
        + token
    )

    conn = common_request.CommonRequest("GET", url)

    if conn.Request():
        print("Success:", conn)
        data = json.loads(conn.ResponseData())
        conn.Close()
        return data

    print("Failure:", conn)
    conn.Close()
    return None



def SaveQuoteData(ticker, jDate, quoteData):


    quoteData["_id"] = ticker.upper() + ":" + jDate
    quoteData["key"] = ticker.upper() + ":" + jDate
    quoteData["symbol"] = ticker.upper()
    quoteData["julian"] = jDate

    print("Saving {}:{}:{}".format(ticker.upper(),jDate,quoteData))

    url = "http://localhost:5984/quotes/" + ticker.upper() + ":" + jDate

    conn = common_request.CommonRequest("PUT", url, myHeader(), quoteData)
    if conn.Request("PUT", url, myHeader(), quoteData):
        print("Successfully inserted", conn)

    else:
        print("Failed:", conn)

def CacheLookup(ticker, jDate = None):

    if jDate == None:
        tt = time.localtime()
        jDate = str(tt.tm_year) + str(tt.tm_yday)


    url = "http://localhost:5984/quotes/" + ticker.upper() + ":" + jDate

    # hdr = myHeader()
    # print(hdr)

    conn = common_request.CommonRequest("GET", url, myHeader())

    if conn.Request():
        print("Success:", conn)
        return json.loads(conn.ResponseData())

    print("Failed Cache Lookup:", conn)

    myData = GetStockQuote(ticker)
    if isinstance(myData,dict):
        SaveQuoteData(ticker,jDate,myData)
        return myData
    else:
        print("Failed GetStockQuote: ", str(myData))

    return None




if __name__ == "__main__":

    ticker = "JNJ"
    tt = time.localtime()
    jDate = str(tt.tm_year) + str(tt.tm_yday)

    if CacheLookup(ticker, jDate) != None:
        #
        print("Yeah!",CacheLookup(ticker,jDate))
        sys.exit(0)

    print("Going to go get it")
    quoteData = GetStockQuote(ticker)
    if ticker != None and isinstance(quoteData,dict):
        print("Saving Quote Data",quoteData)
        SaveQuoteData(ticker,jDate,quoteData)
    else:
        print("what?",quoteData)
