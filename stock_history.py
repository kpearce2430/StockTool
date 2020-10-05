import transaction  # as Transactions
import basedatacsv as bdata
import stocktool

import xlsxwriter
import common_xls_formats
import sys

# import pprint
import datetime
import calendar
import stock_cache


#
#
class HistoryMatrix:
    def __init__(self, wrkbk, fmt, number=36, type="months"):
        self.symbolMatrix = dict()
        self.symbols = []
        self.unitNumber = number
        self.unitType = type
        self.workbook = wrkbk
        self.formats = fmt
        self.columns = {}  # for ColumnInfo class

    def __str__(self):
        output = "unitNumber[{}] unitType[{}]".format(self.unitNumber, self.unitType)

        for k in sorted(self.symbols):
            output = k + "\n"
            s = self.symbolMatrix.get(k)
            output = output + "\n" + str(s) + "\n"

        return output

    def setShares(self, symbol, entry, shares=0.00, price=0.00):

        hdata = self.symbolMatrix.get(symbol)
        if hdata is None:
            #
            hdata = HistoryRows(symbol, self.unitNumber)
            self.symbolMatrix[symbol] = hdata
            self.symbols.append(symbol)

        hdata.addToRow(entry, shares, price)

    def printHistroyMatrix(self):
        print(str(self))

    def writeMatrixWorksheet(self, type = "Quantity", startRow=0, startColumn=0, worksheet=None):

        if worksheet == None:
            if self.workbook == None:
                return None
            myWorksheet = self.workbook.add_worksheet(type + " Matrix")
        else:
            myWorksheet = worksheet

        row = startRow
        column = startColumn

        # 1 Column is the symbol
        ci = common_xls_formats.ColumnInfo(myWorksheet, "Symbol", column)
        self.columns[0] = ci
        ci.columnWrite(row, column, "Symbol", "text", self.formats.headerFormat(), True)
        column = column + 1

        # now write out the months
        for i in range(0, self.unitNumber):
            colDateT = transaction.monthdelta(datetime.date.today(), -i)

            # print(colDateT)
            if calendar.month_abbr[colDateT.month] == "Jan":
                colHeader = ( calendar.month_abbr[colDateT.month] + "/" + str(colDateT.year) )
            else:
                colHeader = calendar.month_abbr[colDateT.month]

            ci = common_xls_formats.ColumnInfo(myWorksheet, colHeader, column)
            self.columns[i] = ci
            ci.columnWrite( row, column, colHeader, "text", self.formats.headerFormat(), True )
            column = column + 1

        row = row + 1
        for key in sorted(self.symbols):
            column = startColumn
            ci.columnWrite(row, column, key, "text", self.formats.textFormat(row))

            column = column + 1
            hdata = self.symbolMatrix.get(key)
            for i in range(0, self.unitNumber):
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

                ci = self.columns[i]
                ci.columnWrite(
                    row, column, fValue, "number", self.formats.numberFormat(row)
                )
                column = column + 1

            # go
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
            print("Something is there, but its not a float[{}]".format(Value))
            return None

    return Value


def findClosingPrice(ticker, theDate=None):
    sPrice = None
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
                print("iex_history close[{}]".format(sPrice))

            elif isinstance(iexQuote, dict):
                sPrice = iexQuote.get("close")  #
                print("iex_quote close[{}]".format(sPrice))

            if sPrice == None:
                iexQuote = priceData.get("quote")
                if isinstance(iexQuote, dict):
                    sPrice = iexQuote.get("close")
                    if sPrice == None or sPrice == "None":
                        print("ERROR - Unable to get price from:()")

    print("sPrince[{}]".format(sPrice))
    price = convertFloat(sPrice)
    if price is None:
        print("WARNING! On {} Price for {} is 'None'".format(theDate,ticker))
        return 0.00

    return price
    # return convertFloat(sPrice)


if __name__ == "__main__":

    # load the lookups
    lookUps = dict()
    stocktool.LoadLookup("lookup.csv", lookUps)

    # intialize the transactions
    transactions = []
    inFilename = "transactions.csv"
    outFilename = "transactions.xlsx"
    cache = stock_cache.StockCache()

    i = 0
    for i in range(1, len(sys.argv)):

        if i == 1:
            inFilename = sys.argv[i]
        elif i == 2:
            outFilename = sys.argv[i]
        else:
            print("Ignoring extra arguments", sys.argv[i])

    # initialize the workbook
    workbook = xlsxwriter.Workbook(outFilename)

    # need the formats
    formats = common_xls_formats.XLSFormats(workbook)

    # get the transacations
    T = transaction.Transactions(workbook, formats)
    T.loadTransactions(inFilename, lookUps)

    # set up the matrix
    hMatrix = HistoryMatrix(workbook, formats, 36, "months" )

    symbols = dict()
    unique_accounts = []

    lastDiff = 999
    for row in T.transactions:
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
                for key, value in symbols.items():
                    if hasattr(value, "numShares"):
                        print(key, mDate)
                        if value.numShares() > 0:
                            # print("{} : shares {} ".format(key, value.numShares()))
                            hMatrix.setShares(key, lastDiff, value.numShares(),findClosingPrice(key,mDate))
                            # print("{} :${}".format(key, findClosingPrice(key,mDate)))
                            #
                    # else:
                    #    print("{} : {} ".format(key, value))

                # bdata.printSymbols(symbols)
            # break

        lastDiff = myDiff
        bdata.ProcessEntry(e, symbols, unique_accounts)

    print("working on last {}".format(lastDiff))

    # warn of gap for now
    # TODO:  If lastdiff isn't zero,there is a gap in the transactions
    if lastDiff > 0:
        print("WARNING: GAP In Transactions with lastDiff[{}]",lastDiff)

    for key, value in symbols.items():
        if hasattr(value, "numShares"):
            if value.numShares() > 0:
                #
                hMatrix.setShares(key, lastDiff, value.numShares(), findClosingPrice(key))
                # print("{} :${}".format(key, findClosingPrice(key)))
                #
       #  else:
       #      print("{} : {} ".format(key, value))

    #
    #
    hMatrix.printHistroyMatrix()
    hMatrix.writeMatrixWorksheet("Quantity")
    hMatrix.writeMatrixWorksheet("Price")
    hMatrix.writeMatrixWorksheet("Value")
    workbook.close()
