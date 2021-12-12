#
# Copyright (c) 2018 Pearce Software Solutions. All rights reserved.
#
# import sys
import csv
import xlsxwriter
import common_xls_formats
import argparse
import datetime
import stock_cache

#
# This class loads the portfolio value output from Quicken as well as
# the lookups.csv.
#  The lookups.csv is a self maintained list of Names and their Symbols.  It is also
#  Used with transactions.py  to identify any stocks that are no longer considered to be relavent.
#  The stocks typcially fall into 2 categories:
#  1. Stocks no longer valid due to $bankruptcy or merger.
#  2. Mutual Funds which do not have a common symbol such as private funds in a 401k
#
#
class PortfolioValue:

    lookups = None
    filename = None
    currentHeaders = []
    #  These are the headers from Quicken with the labels (or tags) that I've assigned them.
    knownFields = {
        "Name": "name",
        "Symbol": "symbol",
        "Type": "type",
        "Price": "quote",
        "Quote": "quote",
        "Price Day Change": "price_day_change",
        "Price Day Change (%)": "price_day_change_pct",
        "Shares": "shares",
        "Cost Basis": "cost_basis",
        "Market Value": "market_value",
        "Average Cost Per Share": "avg_cost_per_share",
        "Gain/Loss 12-Month": "gain_loss_last_12m",
        "Gain/Loss": "gain_loss",
        "Gain/Loss (%)": "gain_loss_pct",
    }

    def __init__(self, portfoliofilename=None, lookupfilename=None):
        #
        self.data = dict()
        self.portfolioFilename = portfoliofilename
        self.lookupFilename = lookupfilename
        self.headers = []
        self.labels = []
        print("PV:{} LU:{}".format(portfoliofilename, lookupfilename))

        if self.lookupFilename != None:
            self.LoadLookup(lookupfilename)

        if self.portfolioFilename != None:
            self.LoadValues(portfoliofilename)

    def __str__(self):
        myStr = (
            self.filename
            + "\ncreated:"
            + str(self.created)
            + "\nheaders:"
            + str(self.headers)
            + "\nlabels:"
            + str(self.labels)
            + "\nlooksups:"
            + str(self.lookups)
            + "\ndata:"
            + str(self.data)
            + "."
        )
        return myStr

    def Headers(self):
        return self.headers

    def Labels(self):
        return self.labels

    def GetValue(self, key, label):
        #
        d = self.data.get(key)
        if d != None:
            #
            value = d.get(label)
            # print("GetValue(",key,":",label,")=",value)
            return value
        else:
            return None

    #
    def LoadValues(self, filename=None):

        # print("In pv.LoadValues")
        if self.filename == None and filename == None:
            print("WARNING: No filename provided to Load Values")
            return

        if filename != None:
            self.filename = filename

        if self.lookups == None:
            print("Error: No lookups provided")
            return

        if self.lookups != None and isinstance(self.lookups, dict) == False:
            # self.lookups = lookups
            print("ERROR: Lookups not valid type")
            return

        foundHeader = False
        numFields = 2
        pvalueReader = csv.reader(
            open(self.filename, newline=""), delimiter=",", quotechar='"'
        )
        i = 0
        self.headers.clear()
        self.labels.clear()
        for row in pvalueReader:
            i = i + 1

            if foundHeader == False and len(row) == 1:
                parts = row[0].split(":")
                if parts[0] == "Created":
                    print(">>", parts)
                    date = parts[1].lstrip()
                    self.created = datetime.datetime.strptime(date, "%Y-%m-%d").date()
                    print("Created:{}".format(date))

                continue

            if len(row) < numFields:
                # print("Skipping row ",i)
                # print(">>>", row)
                continue

            sname = row[0]
            row[0] = sname[1:]
            sname = row[0]

            if sname == "Cash":
                print(sname, ":", row[7])
                continue

            if sname == "Totals":
                print(sname, ":", row[7])
                continue

            # print("row:",i,":",row[0])

            if foundHeader == False:
                #
                #  Assuming 'Name' and 'Symbol' are always the first two columns.
                #
                if len(row) > numFields and row[1] == "Symbol":
                    # Yes, we have the header row
                    row[0] = "Name"

                    for r in row:
                        lbl = self.knownFields.get(r)
                        if lbl != None:
                            self.headers.append(r)
                            self.labels.append(lbl)
                        else:
                            print("Warning: No label/header for {}".format(r))

                    foundHeader = True
                    print("Found Header")
                    numFields = len(row)
                else:
                    print(">>>", row)

                # print("{} : {} : {}".format(len(r),len(self.labels),len(self.headers)))

            else:  # foundHeader = true
                # print("Processing row:",row)
                myValues = dict()
                mySymbol = None
                #
                for j in range(len(self.labels)):
                    #
                    value = row[j]
                    key = self.labels[j]
                    #
                    if key != None:
                        value = value.replace(",", "")
                        value = value.replace("$", "")
                        value = value.replace("#", "")
                        value = value.replace("%", "")

                        if key == "name":
                            myName = value

                        #
                        if key == "symbol":
                            mySymbol = value

                        myValues[key] = value

                if mySymbol == None or len(mySymbol) == 0:
                    if self.lookups != None:
                        mySymbol = self.lookups.get(myName)
                        if mySymbol == None:
                            print("WARNING: Empty mySymbol")
                            mySymbol = myName[0:5]
                            # print(mySymbol)
                    else:
                        mySymbol = myName[0:5]
                        # print("New Symbol",mySymbol)

                    myValues["symbol"] = mySymbol

                # print(mySymbol,":",myValues)
                self.data[mySymbol] = myValues
                # print(self.data[mySymbol])

        print("Loaded {} values".format(i))

    #
    def WriteValuesWorksheet(self, workbook, formats):

        worksheet = workbook.add_worksheet("Portfolio")

        myKeys = self.data.keys()

        columnInfo = []

        myRow = 0
        myColumn = 0
        for h in self.headers:
            ci = common_xls_formats.ColumnInfo(worksheet, h, myColumn)
            ci.columnWrite(myRow, myColumn, h, "text", formats.headerFormat(), True)
            columnInfo.append(ci)
            myColumn = myColumn + 1

        myRow = 1
        for k in myKeys:
            v = self.data[k]
            myColumn = 0
            for l in self.labels:
                ci = columnInfo[myColumn]

                if l == "shares":
                    ci.columnWrite(
                        myRow, myColumn, v[l], "number", formats.numberFormat(myRow)
                    )

                elif (
                    l == "quote"
                    or l == "price_day_change"
                    or l == "market_value"
                    or l == "gain_loss"
                    or l == "avg_cost_per_share"
                    or l == "gain_loss_last_12m"
                    or l == "cost_basis"
                ):
                    ci.columnWrite(
                        myRow, myColumn, v[l], "currency", formats.currencyFormat(myRow)
                    )

                elif l == "price_day_change_pct" or l == "gain_loss_pct":
                    # since I'm already reading in a percent, it needs to be converted back to a float
                    fValue = ci.convertFloat(v[l])
                    if fValue != None:
                        fValue = fValue / 100
                        ci.columnWrite(
                            myRow,
                            myColumn,
                            fValue,
                            "percent",
                            formats.percentFormat(myRow),
                        )
                    else:
                        ci.columnWrite(
                            myRow, myColumn, v[l], "text", formats.textFormat(myRow)
                        )

                else:
                    ci.columnWrite(
                        myRow, myColumn, v[l], "text", formats.textFormat(myRow)
                    )

                myColumn = myColumn + 1

            myRow = myRow + 1

        for ci in columnInfo:
            ci.columnSetSize(1.3)

    def LoadLookup(self, lookupFilename="lookup.csv"):

        # if we're calling this multiple times, basically create
        # a new json struct.
        self.lookups = {}

        lookupReader = csv.reader(
            open(lookupFilename, newline=""), delimiter=",", quotechar='"'
        )
        i = 0
        for row in lookupReader:
            i = i + 1
            if len(row) == 2:
                # print(row[0],row[1])
                self.lookups[row[0]] = row[1]
            else:
                print("huh:", row)

        print("Loaded {} lookups".format(i))

    def WriteLookupWorksheet(self, workbook, formats, startRow=0, startCol=0):

        myRow = startRow
        myColumn = startCol

        worksheet = workbook.add_worksheet("Lookups")

        keyName = "Lookup Key"
        ciKey = common_xls_formats.ColumnInfo(worksheet, keyName, myColumn, 1, 1, 100)
        ciKey.columnWrite(myRow, myColumn, keyName, "text", formats.headerFormat())

        valueName = "Lookup Value"
        ciValue = common_xls_formats.ColumnInfo(
            worksheet, valueName, myColumn + 1, 1, 1, 100
        )
        ciValue.columnWrite(
            myRow, myColumn + 1, valueName, "text", formats.headerFormat(), True
        )

        myRow = myRow + 1
        myKeys = self.lookups.keys()
        for k in sorted(myKeys):
            v = self.lookups[k]
            ciKey.columnWrite(myRow, myColumn, k, "text", formats.textFormat(myRow))
            ciValue.columnWrite(
                myRow, myColumn + 1, v, "text", formats.textFormat(myRow)
            )
            myRow = myRow + 1

        ciKey.columnSetSize(1.4)
        ciValue.columnSetSize(1.4)

    def Lookups(self):

        if hasattr(self, "lookups") and isinstance(self.lookups, dict):
            return self.lookups

        return None


if __name__ == "__main__":

    cache = stock_cache.StockCache()

    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--input", "-i", help="Input CSV File", default="portfolio_value.csv"
    )
    parser.add_argument(
        "--output", "-o", help="Output XLSX File", default="portfolio_value.xlsx"
    )
    parser.add_argument(
        "--lookup",
        "-l",
        help="File containing lookups for translations",
        default="lookup.csv",
    )
    args = parser.parse_args()

    # print(args)

    workbook = xlsxwriter.Workbook(args.output)
    formats = common_xls_formats.XLSFormats(workbook, 5)

    # load up the lookup table
    pv = PortfolioValue(args.input, args.lookup)
    pv.WriteValuesWorksheet(workbook, formats)
    pv.WriteLookupWorksheet(workbook, formats)

    if hasattr(pv, "created") and cache.isCouchDBUp() != None:
        cache.LoadCacheFromJson("portfolio_value", pv.data, date=pv.created)

    print(str(pv))

    workbook.close()
