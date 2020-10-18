#
# Copyright (c) 2020 Pearce Software Solutions. All rights reserved.
#
import transaction  # as Transactions
import basedatacsv as bdata
import argparse

import xlsxwriter
import common_xls_formats
import json

#
import datetime
import calendar
import stock_cache

#
#
#
class HistoryMatrix:

    def __init__(self, wrkbk, fmt, number=36, type="months"):
        self.symbolMatrix = dict() # This is the main data
        self.symbols = dict() # This dict is used for processing transaction
        self.unique_accounts = []
        self.unitNumber = number
        self.unitType = type
        self.workbook = wrkbk
        self.formats = fmt
        self.columns = {}  # for ColumnInfo class

    def __str__(self):
        output = "unitNumber[{}] unitType[{}]".format(self.unitNumber, self.unitType)

        for k in self.symbolMatrix:
            output = k + "\n"
            s = self.symbolMatrix.get(k)
            output = output + "\n" + str(s) + "\n"

        return output

    def setShares(self, symbol, entry, shares=0.00, price=0.00):

        if symbol == None or symbol == "":
            print("WARNING: Invalid Symbol [{}] for Entry: {}".format(symbol,entry))
            raise

        hdata = self.symbolMatrix.get(symbol)
        if hdata is None:
            #
            hdata = HistoryRows(symbol, self.unitNumber)
            self.symbolMatrix[symbol] = hdata
            # self.symbols.append(symbol)
            # print("Added [{}]".format(symbol))

        hdata.addToRow(entry, shares, price)

    def CreateHistoryMatrix(self,transactions):
        self.symbols.clear()
        self.unique_accounts.clear()

        lastDiff = 999
        for row in transactions:
            transDate = row.getDateTimeDate()
            # print()
            # print("{} : {} : {} ".format(row.get_value("security"), row.get_value("date"),row.getTimeTime() ))
            myDiff = monthsDiff(transDate)

            amt = row.getAmount()
            e = bdata.Entry(
                row.get_value("date"),
                row.get_value("type"),
                row.get_value("security_payee"),
                row.get_value("security"),
                row.get_value("description"),
                row.get_value("shares"),
                amt,
                row.get_value("account"),
            )

            if myDiff < 36:

                if myDiff != lastDiff:
                    mDate = monthsAgo(lastDiff, 0)
                    # print("working on {} for {}".format(lastDiff, str(mDate)))
                    for key, value in self.symbols.items():
                        if key == None or key == "":
                            print("Warning, bad key:{}:{}:{}".format(mDate, key, value.numShares()))
                            continue

                        if hasattr(value, "numShares"):
                            print("{}:{}:{}".format(mDate, key, value.numShares()))
                            hMatrix.setShares(key, lastDiff, value.numShares(), findClosingPrice(key, mDate))
                            #

            lastDiff = myDiff
            bdata.ProcessEntry(e, self.symbols, self.unique_accounts)

        print("working on last {}".format(lastDiff))
        print("symbols:\n{}".format(self.symbols))
        # print(json.dumps(symbols),sort_keys=True,indent=4)
        # warn of gap for now
        # TODO:  If lastdiff isn't zero,there is a gap in the transactions
        if lastDiff > 0:
            print("WARNING: GAP In Transactions with lastDiff[{}]", lastDiff)

        for key, value in self.symbols.items():
            if hasattr(value, "numShares"):
                # if value.numShares() > 0:
                    #
                    print("{},{}:{}".format(lastDiff, key, mDate))
                    hMatrix.setShares(key, lastDiff, value.numShares(), findClosingPrice(key))
                    # print("{} :${}".format(key, findClosingPrice(key)))
                    #
        #  else:
        #      print("{} : {} ".format(key, value))

    def printHistroyMatrix(self):
        print(str(self))

    def writeMatrixWorksheet(self, type = "Quantity", startRow=0, startColumn=0, worksheet=None):

        if worksheet == None:
            if self.workbook == None:
                return None
            myWorksheet = self.workbook.add_worksheet(type + " Matrix")
        else:
            myWorksheet = worksheet

        symbolList = []
        for s,hRow in self.symbolMatrix.items():

            if s == "FCASH":
                continue

            if hRow.AnyShares():
                # print("{}:{}".format(s,hRow))
                symbolList.append(s)

        symbolList = sorted(symbolList)

        row = startRow
        column = startColumn
        i = 0
        # 1st Column is the Month
        ci = common_xls_formats.ColumnInfo(myWorksheet, "Month", column)
        self.columns[i] = ci
        ci.columnWrite(row, column, "Symbol", "text", self.formats.headerFormat(), True)
        column = column + 1

        # now write out the symbol column header
        for s in symbolList:
            ci = common_xls_formats.ColumnInfo(myWorksheet, s, column)
            self.columns[i] = ci
            ci.columnWrite( row, column, s, "text", self.formats.headerFormat(), True )
            column = column + 1
            i = i + 1

        # now write out the months on each row
        column = startColumn
        row = row + 1
        ci = self.columns[0]
        # for i in range(0, self.unitNumber):
        for i in range(self.unitNumber,-1,-1):
            rowDateT = transaction.monthdelta(datetime.date.today(), -i)

            # print(colDateT)
            if calendar.month_abbr[rowDateT.month] == "Jan":
                rowHeader = ( calendar.month_abbr[rowDateT.month] + "/" + str(rowDateT.year) )
            else:
                rowHeader = calendar.month_abbr[rowDateT.month]

            ci.columnWrite( row, column, rowHeader, "text", self.formats.textFormat(row), True )
            row = row + 1

        column = startColumn + 1
        for i in range(0,len(self.columns)):
            ci = self.columns[i]
            print(ci.name)

            row = startRow+1
            hdata = self.symbolMatrix.get(ci.name)
            for i in range(self.unitNumber, -1, -1):
            # for i in range(0, self.unitNumber):

                # print("{}:{}".format(i,hdata.months[i]))
                if type == "Quantity":
                    fValue = hdata.Quantity(i)
                elif type == "Price":
                    fValue = hdata.Price(i)
                elif type == "Value":
                    fValue = hdata.Value(i)
                else:
                    print("WARNING: Invalid Type [{}]",type)
                    raise

                if type == "Quantity":
                    ci.columnWrite( row, column, fValue, "number", self.formats.numberFormat(row) )
                else:
                    ci.columnWrite(row, column, fValue, "number", self.formats.currencyFormat(row))
                row = row + 1

            column = column + 1
            # go

        if type == "Value":
            # add a total formula
            ci = common_xls_formats.ColumnInfo(myWorksheet, "Total", column)
            self.columns[column] = ci
            row = startRow
            ci.columnWrite(row, column, "Total", "text", self.formats.headerFormat(), True)
            row = row + 1
            for i in range(0, self.unitNumber+1):
                startSumColumn = xlsxwriter.utility.xl_col_to_name(startColumn+1) + str(row+1)
                endSumColumn = xlsxwriter.utility.xl_col_to_name(column-1) + str(row+1)
                formula = "=SUM("+ startSumColumn + ":" + endSumColumn + ")"
                ci.columnWrite(row,column,formula,"formula",formats.currencyFormat(row))
                row = row + 1

        for myCol in self.columns:
            ci = self.columns[myCol]
            ci.columnSetSize(1.3)


#
# The symbol data for month, week and day.
#
class HistoryRows:
    def __init__(self, s=None, r=36):
        self.symbol = s
        # self.unitType = t # month, week, day
        self.unitQuantity = [0.00] * (r + 1)
        self.unitPrice = [0.00] * (r + 1)

    def __str__(self):

        valueStr = ""
        for i in range(0, len(self.unitQuantity) - 2):
            valueStr = valueStr + str(self.Value(i)) + ","

        valueStr = valueStr + str(self.Value(len(self.unitQuantity) - 1))
        return (
            self.symbol.upper()
            + "\n\tunitQuantity:"
            + str(self.unitQuantity)
            + "\n\tunitPrice"
            + str(self.unitPrice)
            + "\n\tValue["
            + valueStr
            + "]"
        )

    def addToRow(self, entry = 0, quantity=0.00, price=0.00):

        if entry < len(self.unitQuantity):
            if isinstance(quantity, float):
                self.unitQuantity[entry] = quantity
            else:
                print("ERROR: Quantity [{}] required to be float".format(quantity))
                raise

            if isinstance(price, float):
                self.unitPrice[entry] = price
            else:
                print("ERROR: Price [{}] required to be float".format(price))
                raise

        else:
            print( "Error entry {} outside range of months {}".format( entry, len(self.unitQuantity) ) )
            raise

    def AnyShares(self):
        for i in range(0,len(self.unitQuantity)):
            if self.unitQuantity[i] > 0.00:
                return True

        return False

    def Quantity(self, entry):
        if entry < len(self.unitQuantity):
            return self.unitQuantity[entry]

        print( "ERROR: {} outside range of quantity {}".format(entry,len(self.unitQuantity)))
        raise

    def Price(self, entry):
        if entry < len(self.unitPrice):
            return self.unitPrice[entry]

        print("ERROR: {} outside range of price {}".format(entry,len(self.unitPrice)))
        raise

    def Value(self, entry):
        if entry < len(self.unitQuantity):
            return self.unitQuantity[entry] * self.unitPrice[entry]

        print("ERROR: {} outside range of price {}".format(entry,len(self.unitPrice)))
        raise


def monthsAgo(mon, day=1, d=None):

    if d == None:
        d1 = datetime.date.today()
    elif hasattr(d, "year") and hasattr(d, "month"):
        d1 = d
    else:
        print("Error, invalid data object {} ".format(d))
        return None

    if mon > 11:
        yAgo = int(mon / 12)
        mAgo = int(mon % 12)
    else:
        yAgo = 0
        mAgo = mon

    dYear = d1.year - yAgo
    dMonth = d1.month - mAgo

    if dMonth <= 0:
        dYear = dYear - 1
        dMonth = dMonth + 12

    # Returns the last day of the month
    if day <= 0:
        days = calendar.monthrange(dYear, dMonth)
        day = days[1]

    return datetime.date(dYear, dMonth, day)


def monthsDiff(date1, date2=None):

    if date2 == None:
        d2 = date1
        d1 = datetime.date.today()
    else:
        d1 = date1
        d2 = date2

    return (d1.year - d2.year) * 12 + d1.month - d2.month


def monthsAbs(date1, date2=datetime.date.today()):
    return abs(monthsDiff(date1, date2))


def convertFloat(Value):
    if isinstance(Value, float) == False:
        try:
            fValue = float(Value)
            return fValue
        except:
            # write nothing
            print("WARNING: Something is there, but its not a float[{}]".format(Value))
            return 0.00
            # return None

    return Value


def findClosingPrice(ticker, theDate=None):
    sPrice = None
    if ticker == "CTL" or ticker == "FCASH" or ticker == "MMDA1":
        # TODO: Some bug around Century Link
        return 1.00

    if len(ticker) <= 4:
        if theDate == None:
            priceData = cache.StockLookup(ticker)
        else:
            priceData = cache.StockLookup(ticker, time=theDate.timetuple())

        if isinstance(priceData, dict):
            # print("Lookup Returns: {}:{}".format(key, priceData))
            iexHistory = priceData.get("iex_history")
            iexQuote = priceData.get("iex_quote")
            if isinstance(iexHistory, dict):
                # unchanged Close since we have unchanged number of shares
                sPrice = iexHistory.get("uClose") #
                # print("iex_history close[{}]".format(sPrice))

            elif isinstance(iexQuote, dict):
                sPrice = iexQuote.get("close")  #
                # print("iex_quote close[{}]".format(sPrice))

            if sPrice == None:
                iexQuote = priceData.get("quote")
                if isinstance(iexQuote, dict):
                    sPrice = iexQuote.get("close")
                    if sPrice == None or sPrice == "None":
                        print("ERROR - Unable to get price from:()")
    else:
        # print("Mutual Fund [{}] Date [{}]".format(ticker,str(theDate)))
        if theDate == None:
            # print("{} for today".format(ticker))
            priceData = cache.MutualFundsLookup(ticker)
        else:
            priceData = cache.MutualFundsLookup(ticker,time=theDate.timetuple())

        if isinstance(priceData,dict):
            print("{} found price data {}".format(ticker,priceData))
            portfolioValue = priceData.get("portfolio_value")
            fundsData = priceData.get("fund_record")
            transData = priceData.get("transaction")

            if isinstance(fundsData,dict):
                sPrice = fundsData.get("close")
                print("Found price {} in fundsData".format(sPrice))

            if sPrice == None and isinstance(portfolioValue,dict):
                sPrice = portfolioValue.get("quote")
                print("Found price {} in portfolioValue".format(sPrice))

            if sPrice == None and isinstance(transData,dict):
                sPrice = transData.get("quote")
                print("Found price {} in transaction".format(sPrice))

            if theDate  != None:
                print("{}:{} Hit with sPrice[{}]".format(ticker,cache.jDateFromTime(theDate.timetuple()),sPrice))
            else:
                print("{}:Today Hit with sPrice[{}]".format(ticker,sPrice))

        elif theDate == None:
            print("WARNING:  No Data for {} for today in cache".format(ticker))
            sPrice="0.00"
        else:
            print("WARNING:  No Data for {} for {} in cache".format(ticker,theDate))
            jDate = cache.jDateFromTime(theDate.timetuple())

            selectorOperator = {}
            selectorOperator["$lte"] = ticker+":"+jDate
            selectorFields = {}
            selectorFields ["key"] = selectorOperator
            selectorData = {}
            selectorData["selector"] = selectorFields
            selectorSortFields = {}
            selectorSortFields["key"] = "desc"
            selectorSort = []
            selectorSort.append(selectorSortFields)
            selectorData["sort"] = selectorSort
            selectorData["limit"] = 5

            print("{}:{}:{}".format(ticker,jDate,selectorData))
            response = cache.couchFindByPartition(ticker,"funds",selectorData)
            print("response:{}".format(response))

            if not isinstance(response,dict):
                print("ERROR Response form couchFindByPartition")
                raise
                # sys.exit(0)

            docs = response.get("docs")
            if not isinstance(docs,list):
                print("ERROR: Expected docs as a list and got {}",docs)
                raise
                # sys.exit(0)

            if len(docs) <= 0:
                print("No docs found for {} on {}",ticker,jDate)
                return 0.00

            for d in docs:
                if not isinstance(d,dict):
                    print("ERROR Expecing a json blob not {}".format(d))
                    raise

                pv = d.get("portfolio_value")
                if isinstance(pv,dict):
                    sPrice = pv.get("quote")
                    print("portfolio value quote for {} [{}]".format(ticker,sPrice))

                if sPrice == None:
                    tr = d.get("transaction")
                    if isinstance(tr,dict):
                        sPrice = tr.get("quote")
                        print("transaction quote for {} [{}]".format(ticker,sPrice))

                if sPrice == None:
                    print("WARNING: No price found for {} on {}",ticker,jDate)
                else:
                    break


    # print("sPrince[{}]".format(sPrice))
    price = convertFloat(sPrice)
    if price is None:
        print("WARNING! On {} Price for {} is 'None'".format(theDate,ticker))
        raise
        # return 0.00

    return price
    # return convertFloat(sPrice)


if __name__ == "__main__":

    # load the lookups
    parser = argparse.ArgumentParser()
    parser.add_argument( "--input", "-i", help="Input CSV File", default="transactions.csv" )
    parser.add_argument( "--output", "-o", help="Output XLSX File", default="stock_history.xlsx" )
    parser.add_argument( "--lookup", "-l", help="File containing lookups for translations", default="lookup.csv" )
    parser.add_argument( "--portfolio", "-p", help="Portfolio Values Lookups", default="portfolio_value.csv" )
    args = parser.parse_args()

    import portfoliovalue
    pv = portfoliovalue.PortfolioValue(args.portfolio, args.lookup)

    # intialize the transactions
    transactions = []
    cache = stock_cache.StockCache()

    # initialize the workbook
    workbook = xlsxwriter.Workbook(args.output)

    # need the formats
    formats = common_xls_formats.XLSFormats(workbook)

    # get the transacations
    T = transaction.Transactions(workbook, formats)
    T.loadTransactions(args.input, pv.lookups)

    # set up the matrix
    hMatrix = HistoryMatrix(workbook, formats, 36, "months" )
    hMatrix.CreateHistoryMatrix(T.transactions)
    hMatrix.printHistroyMatrix()
    hMatrix.writeMatrixWorksheet("Quantity")
    hMatrix.writeMatrixWorksheet("Price")
    hMatrix.writeMatrixWorksheet("Value")
    workbook.close()



