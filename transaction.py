#
# Copyright (c) 2018 Pearce Software Solutions. All rights reserved.
#
import sys
import csv
import xlsxwriter
import json
import copy
import common_xls_formats
import datetime
import time
import numpy
import calendar


class Transaction:
    def __init__(self, entry=None):
        # self.name = ""
        self.info = dict()

        if isinstance(entry, dict):
            for k in entry:
                self.info[k] = entry[k]

    def __str__(self):
        return "Transaction:" + json.dumps(self.info)

    def set_value(self, key=None, value=None):

        if key != None:
            self.info[key] = value
        else:
            print("Error - Missing key in Transaction.set_value")

    def get_value(self, key):
        if key != None:
            return self.info.get(key)
        else:
            return None

    def getDate(self):
        return self.get_value("date")

    def getDateTimeDate(self):
        if not hasattr(self, "transDatetime"):
            self.transDateTime = datetime.date(
                int(self.getYear()), int(self.getMonth()), int(self.getDay())
            )

        return self.transDateTime

    def getTimeTime(self):
        if not hasattr(self, "transTimeTime"):
            self.transTimeTime = time.strptime(self.getDate(),"%m/%d/%Y")

        return self.transTimeTime


    def getYear(self):
        if not hasattr(self, "year"):
            d = self.getDate()
            ds = d.split("/")
            self.year = ds[2]
            self.month = ds[0]
            self.day = ds[1]

        return self.year

    def getMonth(self):
        if not hasattr(self, "month"):
            d = self.getDate()
            ds = d.split("/")
            self.year = ds[2]
            self.month = ds[0]
            self.day = ds[1]

        return self.month

    def getDay(self):
        if not hasattr(self, "day"):
            d = self.getDate()
            ds = d.split("/")
            self.year = ds[2]
            self.month = ds[0]
            self.day = ds[1]

        return self.day

    def getAmount(self):
        if not hasattr(self, "amount"):

            if self.get_value("type") == "Reinvest Dividend" or self.get_value("type") == "Add Shares":
                self.amount = self.get_value("invest_amt")
            else:
                self.amount = self.get_value("amount")

        return self.amount


class Transactions:
    def __init__(self, workbook, formats):

        self.workbook = workbook
        self.formats = formats
        self.transactions = []
        self.jsonTags = []
        self.headers = []
        self.transColumns = {}  # for ColumnInfo class
        self.divColumns = {}  # for Dividend sheet

    def loadTransactions(self, infilename, lookups):

        if isinstance(lookups, dict) == False:
            print("Invalid argument lookups")
            return

        transReader = csv.reader(
            open(infilename, newline=""), delimiter=",", quotechar='"'
        )

        i = 0
        securityRow = -1
        symbolRow = -1
        myTags = []
        # myHeaders = []

        for row in transReader:

            i = i + 1

            if len(row) < 3:
                continue

            # The first two columns from Quicken 18 are junk.
            # Get rid of them.
            del row[0]
            del row[0]

            # Skip the hard coded Date line.
            if row[0] == "Date" and len(self.headers) == 0 and len(self.jsonTags) == 0:
                j = 0
                for r in row:
                    self.headers.append(r)
                    tag = returnTag(r)
                    self.jsonTags.append(tag)
                    if r == "Security":
                        securityRow = j
                    if r == "Symbol":
                        symbolRow = j

                    j = j + 1

                print("# Tags:", len(self.jsonTags))
                print("# Headers:", len(self.headers))
                print("Security Row:", securityRow)
                print("Symbol Row:", symbolRow)
                continue

            if len(self.jsonTags) <= 0:
                # havent found the header row yet. lets just...
                continue

            if len(row) != len(self.headers):
                continue

            info = Transaction()

            for j in range(len(self.jsonTags)):

                if self.jsonTags[j] == "None":
                    continue

                # Get the symbols from the lookup
                if self.headers[j] == "Security":
                    # I found over the years that symbols occassionally get reused.  I use
                    # lookups to override any symbols that Quicken may put in.  If there isn't
                    # a symbol in lookup and symbols are provided in the sheet.  I use the
                    # sheet.
                    value = lookups.get(row[j])
                    if value == None:
                        if symbolRow != -1:
                            value = row[symbolRow]
                        else:
                            value = "Missing"

                    row[j] = value

                    if row[symbolRow] == None or row[symbolRow] == "":
                        row[symbolRow] = value

                if (
                    self.headers[j] == "Shares"
                    or self.headers[j] == "Amount"
                    or self.headers[j] == "Invest Amt"
                ):
                    # Remove the commas
                    row[j] = row[j].replace(",", "")

                info.set_value(self.jsonTags[j], row[j])
                j = j + 1

            i = i + 1

            self.transactions.append(info)

        self.headers.append("Year")
        self.jsonTags.append("year")
        self.headers.append("Month")
        self.jsonTags.append("month")

        print("Loaded ", len(self.transactions), " transactions")

    def writeTransactionWorksheetHeaders(self, worksheet, startRow=0, startColumn=0):
        self.transColumns.clear()
        myCol = startColumn
        # headers = TransactionHeaders()
        for h in self.headers:
            ci = common_xls_formats.ColumnInfo(worksheet, h, myCol)
            self.transColumns[myCol] = ci
            ci.columnWrite(
                startRow, myCol, h, "text", self.formats.headerFormat(), True
            )
            myCol = myCol + 1

    def writeTransactionWorksheetRow(self, worksheet, trans, startRow=0, startColumn=0):

        myCol = startColumn
        dateCol = xlsxwriter.utility.xl_col_to_name(myCol)

        for i in range(len(self.headers)):

            myTag = self.jsonTags[i]
            value = trans.get_value(myTag)
            ci = self.transColumns[myCol]

            if myTag == "date":
                ci.columnWrite(
                    startRow, myCol, value, "date", self.formats.dateFormat(startRow)
                )

            elif myTag == "shares":
                ci.columnWrite(
                    startRow,
                    myCol,
                    value,
                    "number",
                    self.formats.numberFormat(startRow),
                )

            elif myTag == "amount":
                # if trans.get_value("type") == "Reinvest Dividend":
                value = trans.getAmount()
                ci.columnWrite(
                    startRow,
                    myCol,
                    value,
                    "currency",
                    self.formats.currencyFormat(startRow),
                )

            elif myTag == "invest_amt":
                value = trans.get_value("invest_amt")
                ci.columnWrite(
                    startRow,
                    myCol,
                    value,
                    "currency",
                    self.formats.currencyFormat(startRow),
                )

            elif myTag == "year":
                formula = "=YEAR(" + dateCol + str(startRow + 1) + ")"
                ci.columnWrite(
                    startRow,
                    myCol,
                    formula,
                    "formula",
                    self.formats.formulaFormat(startRow),
                )

            elif myTag == "month":
                formula = "=MONTH(" + dateCol + str(startRow + 1) + ")"
                ci.columnWrite(
                    startRow,
                    myCol,
                    formula,
                    "formula",
                    self.formats.formulaFormat(startRow),
                )

            else:
                ci.columnWrite(
                    startRow, myCol, value, "date", self.formats.textFormat(startRow)
                )

            myCol = myCol + 1

    def writeTransactions(
        self, startRow=0, startColumn=0, worksheet=None, filterFunc=None, fArgs=None
    ):

        if worksheet == None:
            if self.workbook == None:
                return None
            myWorksheet = self.workbook.add_worksheet("Transactions")
        else:
            myWorksheet = worksheet

        row = startRow
        column = startColumn
        self.writeTransactionWorksheetHeaders(myWorksheet, row, column)

        row = row + 1
        for t in self.transactions:
            if filterFunc != None:
                if filterFunc(t, fArgs) == True:
                    # print("writing")
                    self.writeTransactionWorksheetRow(myWorksheet, t, row, startColumn)
                    row = row + 1
            else:
                self.writeTransactionWorksheetRow(myWorksheet, t, row, startColumn)
                row = row + 1

        for myCol in self.transColumns:
            ci = self.transColumns[myCol]
            ci.columnSetSize(1.3)
            # print("column:",myCol,",",str(ci))
            # myWorksheet.set_column(ci.columnNumber, ci.columnNumber, ci.columnSize(1.3))

    def getTransactions(self, filterFunc, *args):
        #
        results = []
        for t in self.transactions:
            if filterFunc != None:
                if filterFunc(t, args) == True:
                    print("match:", t)
                    p = copy.deepcopy(t)
                    # p = t.deepcopy(dict)
                    results.append(p)
                    # self.writeTransactionWorksheetRow(myWorksheet, t, row, startColumn)
                    # row = row + 1
            # else:
            # self.writeTransactionWorksheetRow(myWorksheet, t, row, startColumn)
            # row = row + 1

        return results

    def getDividends(self, monthsAgo=36, startRow=0, startColumn=0):

        #  These are the types of transactions that are needed to fill out
        #  the dividend sheet
        # tTypes = ["Dividend Income", "Reinvest Dividend", "Interest Income"]
        tTypes = ["Dividend Income", "Reinvest Dividend", "Interest Income"]

        pickList = dict()
        pickList["type"] = tTypes
        pickList["months"] = monthsAgo
        # pickList["symbol"] = ["AAPL", "HD","USAIX","T"]

        matchingTrans = []

        for t in self.transactions:

            if pickByFields(t, pickList):
                p = copy.deepcopy(t)
                matchingTrans.append(p)

        # Now we have all the transactions required to make the
        # Dividend sheet.
        # print(len(matchingTrans))

        # Need a list of the symbols
        symbolList = []
        dividendData = {}

        # Dont change this value:
        endDate = datetime.datetime.today()

        for m in matchingTrans:
            sym = m.get_value("symbol")
            try:
                symbolList.index(sym)
            except ValueError:
                symbolList.append(sym)

            sDiv = dividendData.get(sym)
            if sDiv == None:
                sDiv = {}
                sDiv["symbol"] = sym
                sDiv["dividends"] = numpy.zeros((monthsAgo + 1))
                divSheet = sDiv["dividends"]
                dividendData[sym] = sDiv
            else:
                divSheet = sDiv["dividends"]

            transDate = m.getDateTimeDate()
            tranMonthsAgo = (endDate.year - transDate.year) * 12 + (
                endDate.month - transDate.month
            )

            divSheet[tranMonthsAgo] += float(m.getAmount())

        #
        symbolList.sort()

        if self.workbook == None:
            return None

        self.divColumns.clear()  # for Dividend sheet

        myWorksheetName = "Dividends_" + str(monthsAgo)

        myWorksheet = self.workbook.add_worksheet(myWorksheetName)
        myRow = startRow
        myCol = startColumn
        # headers = TransactionHeaders()

        ci = common_xls_formats.ColumnInfo(myWorksheet, "Symbol", myCol)
        self.divColumns[myCol] = ci
        ci.columnWrite(
            startRow, myCol, "Symbol", "text", self.formats.headerFormat(), True
        )
        myCol += 1

        for i in range(0, monthsAgo + 1):

            colDateT = monthdelta(datetime.date.today(), -i)

            colHeader = calendar.month_abbr[
                colDateT.month
            ]  # + "/" + str(colDateT.year)

            ci = common_xls_formats.ColumnInfo(myWorksheet, colHeader, myCol)
            self.divColumns[myCol] = ci
            ci.columnWrite(
                startRow, myCol, colHeader, "text", self.formats.headerFormat(), True
            )
            myCol = myCol + 1

        myRow += 1
        for s in symbolList:
            myCol = 0

            ci.columnWrite(
                myRow, myCol, s, "text", self.formats.textFormat(myRow), True
            )
            myCol += 1
            sDiv = dividendData[s]
            mDivs = sDiv["dividends"]
            for i in range(0, len(mDivs)):
                ci.columnWrite(
                    myRow,
                    myCol,
                    mDivs[i],
                    "accounting",
                    self.formats.accountingFormat(myRow),
                )
                myCol += 1
            myRow += 1

        ci = self.divColumns[0]
        ci.columnWrite(myRow, 0, "Total", "text", self.formats.textFormat(myRow))
        ci.columnSetSize(1.3)

        for myCol in range(1, monthsAgo + 2):
            ci = self.divColumns[myCol]
            cName = xlsxwriter.utility.xl_col_to_name(myCol)
            myFormula = (
                "=SUM(" + cName + str(startRow + 2) + ":" + cName + str(myRow) + ")"
            )
            ci.columnWrite(
                myRow, myCol, myFormula, "formula", self.formats.accountingFormat(myRow)
            )
            ci.columnSetSize(1.3)

        # set up the chart basics:
        myChart = self.workbook.add_chart({"type": "column"})
        myChart.set_title({"name": "Dividend Chart"})
        myChart.set_size({"width": 1000, "height": 700})

        # colors = [ "#FF9900", "#00FF00","#0000FF"]
        colors = ["#4DA6FF", "#88FF4B", "#B30059"]

        numSeries = int(monthsAgo / 12)

        # print("numSeries:", numSeries)

        for i in range(0, numSeries):
            seriesStartCol = 1 + (i * 12)
            columnStart = xlsxwriter.utility.xl_col_to_name(seriesStartCol)

            seriesStopCol = 12 + (i * 12)
            columnEnd = xlsxwriter.utility.xl_col_to_name(seriesStopCol)

            myValues = (
                "="
                + myWorksheetName
                + "!$"
                + columnStart
                + "$"
                + str(myRow + 1)
                + ":$"
                + columnEnd
                + "$"
                + str(myRow + 1)
            )

            myColValues = (
                "="
                + myWorksheetName
                + "!$"
                + columnStart
                + "$"
                + str(startRow + 1)
                + ":$"
                + columnEnd
                + "$"
                + str(startRow + 1)
            )
            # print(myValues)

            myYear = endDate.year - i
            myChart.add_series(
                {
                    "name": str(myYear),
                    "values": myValues,
                    "categories": myColValues,
                    "fill": {"color": colors[i]},
                }
            )

        # Insert the chart into the worksheet.
        colName = xlsxwriter.utility.xl_col_to_name(startColumn+1)
        myWorksheet.insert_chart( colName + str( startRow + 2 ), myChart)


def TransactionJsonTags():
    return [
        "date",
        "type",
        "security",
        "symbol",
        "security_payee",
        "description",
        "shares",
        "invest_amt",
        "amount",
        "account",
        "year",
        "month",
    ]


def TransactionHeaders():
    return [
        "Date",
        "Type",
        "Security",
        "Symbol",
        "Security/Payee",
        "Description/Category",
        "Shares",
        "Invest Amt",
        "Amount",
        "Account",
        "Year",
        "Month",
    ]


def returnTag(hdr):
    headers = TransactionHeaders()
    tags = TransactionJsonTags()

    try:
        idx = headers.index(hdr)
    except:
        print("No element found for:", hdr)
        return "None"

    if idx >= len(tags):
        print("Error:  Tag length mismatch")

    return tags[idx]


def pickSymbol(trans, fArgs):
    mySymbol = trans.get_value("symbol")
    # print(fArgs,mySymbol)
    if mySymbol == fArgs:
        # print(fArgs, mySymbol,"True")
        return True
    else:
        # print(fArgs, mySymbol,"False")
        return False


#
# fieldList is expected to be a set of fields with the Transaction
# where all the individual fields are 'and'.
# if the field is a list, those values are 'or'.
#  In the case of 'date', do the date math to compute if the transaction date is greater than the
#  date in the fieldList
#
def pickByFields(trans, fieldList):

    if not isinstance(fieldList, dict):
        return True

    for k in fieldList.keys():
        v = fieldList[k]

        if k == "date" or k == "months":

            # tDate = trans.get_value("date")
            # tDateV = tDate.split("/")
            # tDateT = datetime.date(int(trans.getYear), int(), int(tDateV[1]))
            tDateT = trans.getDateTimeDate()

            if k == "date":
                kDate = fieldList[k]
                kDateV = kDate.split("/")
                kDateT = datetime.date(int(kDateV[2]), int(kDateV[0]), int(kDateV[1]))
                if tDateT > kDateT:
                    continue
                else:
                    return False
            else:  # months
                monthsAgo = int(v)
                yearsAgo = int(monthsAgo / 12)
                monthsAgo = int(monthsAgo % 12)

                nowDateT = datetime.datetime.now()
                prevDateT = datetime.date(
                    nowDateT.year - yearsAgo, nowDateT.month - monthsAgo, 1
                )  # always to the 1st of the month
                # print("Prev Date:",prevDateT)

                if tDateT > prevDateT:
                    continue
                else:
                    return False

        elif isinstance(v, list):
            foundIt = False
            for l in v:
                if trans.get_value(k) == l:
                    foundIt = True
                    break  # for

            if foundIt == False:  # no matching in list
                return False

        elif trans.get_value(k) != v:
            return False

    # if we made it all the way here, return TRUE
    return True


def pickByDate(trans, *args):

    myArgs = args[0]
    if len(myArgs) < 2:
        print("Nope, need symbol and date:", len(myArgs), ":", myArgs)
        sys.exit(-1)
        return False

    mySymbol = myArgs[0]

    tSymbol = trans.get_value("symbol")
    if tSymbol != mySymbol:
        return False

    myDate = myArgs[1]
    myDateValues = myDate.split("/")
    # datetime.date(2008, 6, 24)
    myDater = datetime.date(
        int(myDateValues[2]), int(myDateValues[0]), int(myDateValues[1])
    )

    tDate = trans.get_value("date")
    myTransValues = tDate.split("/")
    myTransDate = datetime.date(
        int(myTransValues[2]), int(myTransValues[0]), int(myTransValues[1])
    )

    if myTransDate > myDater:
        return True
    else:
        return False


def monthdelta(date, delta):
    m, y = (date.month + delta) % 12, date.year + ((date.month) + delta - 1) // 12
    if not m:
        m = 12
    d = min(
        date.day,
        [
            31,
            29 if y % 4 == 0 and not y % 400 == 0 else 28,
            31,
            30,
            31,
            30,
            31,
            31,
            30,
            31,
            30,
            31,
        ][m - 1],
    )
    return date.replace(day=d, month=m, year=y)


if __name__ == "__main__":
    lookUps = dict()
    transactions = []
    inFilename = "transactions.csv"
    outFilename = "transactions.xlsx"

    i = 0
    for i in range(1, len(sys.argv)):

        if i == 1:
            inFilename = sys.argv[i]
        elif i == 2:
            outFilename = sys.argv[i]
        else:
            print("Ignoring extra arguments", sys.argv[i])

    workbook = xlsxwriter.Workbook(outFilename)
    formats = common_xls_formats.XLSFormats(workbook)
    T = Transactions(workbook, formats)
    T.loadTransactions(inFilename, lookUps)
    T.writeTransactions(0, 0)

    myWorksheet = workbook.add_worksheet("APPLE")
    T.writeTransactions(2, 2, myWorksheet, pickSymbol, "AAPL")

    print("Getting apple transactions")
    arg = ("AAPL", "01/01/2018")
    aaplTrans = T.getTransactions(pickByDate, "AAPL", "01/01/2018")
    print(len(aaplTrans))

    print("Building Dividend Sheet")
    T.getDividends()
    workbook.close()
