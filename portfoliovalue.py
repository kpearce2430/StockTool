#
# Copyright (c) 2018 Pearce Software Solutions. All rights reserved.
#
import sys
import csv
import xlsxwriter
import common_xls_formats


class PortfolioValue:

    lookups = None
    filename = None
    currentHeaders = []
    #  These are the headers from Quicken with the labels (or tags) that I've assigned them.
    knownFields = {
        "Name": "name",
        "Symbol": "symbol",
        "Type": "type",
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

    def __init__(self, filename=None, lookups=None, loadIt=True):
        #
        self.data = dict()
        self.filename = filename
        self.headers = []
        self.labels = []

        if lookups != None and isinstance(lookups, dict):
            self.lookups = lookups

        if loadIt == True and self.filename != None:
            self.LoadValues()

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
    def LoadValues(self, filename=None, lookups=None):

        # print("In pv.LoadValues")
        if self.filename == None and filename == None:
            print("No filename provided")
            return

        if filename != None:
            self.filename = filename

        if self.lookups == None and lookups == None:
            print("No lookups provided")
            # return

        if lookups != None and isinstance(lookups, dict) == False:
            self.lookups = lookups
            print("Lookups not valid type")
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
            if len(row) < numFields:
                # print("Skipping row ",i)
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
                        self.headers.append(r)
                        self.labels.append(lbl)

                if len(r) == len(self.labels) and len(r) == len(self.headers):
                    foundHeader = True
                    numFields = len(row)

            else:  # foundHeader = true
                # print("Processing row:",row)
                myValues = dict()
                mySymbol = None
                # print(self.labels)
                for j in range(len(self.labels)):
                    # print(j)
                    value = row[j]
                    key = self.labels[j]
                    # print(key,":",value)
                    if key != None:
                        value = value.replace(",", "")
                        value = value.replace("$", "")
                        value = value.replace("#", "")
                        value = value.replace("%", "")

                        if key == "name":
                            myName = value
                            # print("myName[", myName, "]")

                        if key == "symbol":
                            mySymbol = value
                            # print("mySymbol[",mySymbol,"]",len(mySymbol))

                        myValues[key] = value

                # print(myName,mySymbol)
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

    #
    def WriteValues(self, workbook, formats):

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


if __name__ == "__main__":

    inFilename = "portfolio_value.csv"
    outFilename = "portfolio_value.xlsx"

    i = 0
    for i in range(1, len(sys.argv)):

        if i == 1:
            inFilename = sys.argv[i]
        elif i == 2:
            outFilename = sys.argv[i]
        else:
            print("Ignoring extra arguments", sys.argv[i])

    workbook = xlsxwriter.Workbook(outFilename)
    formats = common_xls_formats.XLSFormats(workbook, 5)

    pv = PortfolioValue(inFilename)
    pv.WriteValues(workbook, formats)

    workbook.close()
