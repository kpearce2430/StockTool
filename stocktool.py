#
# ----------------------------------------------------------------------------
# "THE BEER-WARE LICENSE" (Revision 42):
# As long as you retain this notice you
# can do whatever you want with this stuff. If we meet some day, and you think
# this stuff is worth it, you can buy me a beer in return.
# ----------------------------------------------------------------------------
#
# Copyright (c) 2018 Pearce Software Solutions. All rights reserved.
#
import sys
import csv
import json
import xlsxwriter
import calendar

import basedatacsv as bdata
import portfoliovalue as pv
import transaction
import urllib3
import datetime
import common_xls_formats
import argparse
import stock_cache
import stock_history
import numpy


def printData(labelData, ticker, formats, row=0, col=0):
    myRow = row
    myColumn = col

    for d in labelData:
        tag = d.tag
        field = ticker.get(tag)
        if field != None:
            worksheet.write(myRow, myColumn, d.label, formats.textFormat(myRow))
            if d.format != None:
                # print(d.tag,",",d.format)
                if d.type == "timestamp":
                    tsInt = int(field)
                    if tsInt > 0:
                        tsDt = datetime.datetime.fromtimestamp(tsInt / 1000.0)
                        worksheet.write(myRow, myColumn + 1, tsDt, d.format(myRow))
                    else:
                        print("Field ", d.tag, " has no timestime value:", field)
                        worksheet.write(
                            myRow, myColumn + 1, field, formats.textFormat(myRow)
                        )
                else:
                    worksheet.write(myRow, myColumn + 1, field, d.format(myRow))
            else:
                print("Missing format for ", d.tag, ",", d.type)
                worksheet.write(myRow, myColumn + 1, field, None)
            myRow = myRow + 1

    return myRow


def printEntries(worksheet, entries, formats, startRow=0, startColumn=0):

    entryColumnInfo = bdata.CreateEntryTypes()
    myColumn = startColumn
    myRow = startRow

    for ec in entryColumnInfo:
        ec.ColumnInfo = common_xls_formats.ColumnInfo(
            worksheet, ec.Header, myColumn, 1, 1
        )
        ec.ColumnInfo.columnWrite(
            myRow, myColumn, ec.Header, "text", formats.headerFormat(), True
        )
        myColumn = myColumn + 1

    myRow = myRow + 1
    for e in entries:
        myColumn = startColumn
        for ec in entryColumnInfo:
            ci = ec.ColumnInfo
            tag = ec.Tag
            value = e.get(tag)
            if ec.Header == "Sold Lots":
                value = len(value)  # should be a list of the lots

            if ec.Type == "currency":
                ci.columnWrite(
                    myRow, myColumn, value, ec.Type, formats.currencyFormat(myRow)
                )
            elif ec.Type == "numeric":
                ci.columnWrite(
                    myRow, myColumn, value, ec.Type, formats.numberFormat(myRow)
                )
            elif ec.Type == "date":
                ci.columnWrite(
                    myRow, myColumn, value, ec.Type, formats.dateFormat(myRow)
                )
                dateCol = xlsxwriter.utility.xl_col_to_name(myColumn)
            elif ec.Type == "formula":
                if ec.Header == "Month":
                    fmt = "=MONTH(" + dateCol + str(myRow + 1) + ")"
                elif ec.Header == "Year":
                    fmt = "=YEAR(" + dateCol + str(myRow + 1) + ")"
                ci.columnWrite(
                    myRow, myColumn, fmt, ec.Type, formats.formulaFormat(myRow)
                )
            else:
                ci.columnWrite(
                    myRow, myColumn, value, ec.Type, formats.formulaFormat(myRow)
                )

            myColumn = myColumn + 1

        myRow = myRow + 1

    #
    for ec in entryColumnInfo:
        ec.ColumnInfo.columnSetSize(1.4)


def printMatrix(
    workbook,
    worksheetName,
    worksheet,
    formats,
    ticker,
    histData,
    divData,
    startRow=0,
    startColumn=0,
):

    # print(str(histData))
    # print("printMatrix: {}".format(divData))
    # Create the column info
    hdrs = ["Stock", "Price", "Value", "Quantity", "Dividends"]
    jTags = ["stock", "price", "value", "qty", "divs"]
    mTypes = ["text", "current", "currency", "number", "currency"]
    myColumn = startColumn
    myRow = startRow

    # Create the column data for the Stock, Price, Value, and Quantity section
    entryTypes = []
    columns = []
    for i in range(0, len(hdrs)):
        et = bdata.EntryType(hdrs[i], jTags[i], mTypes[i])
        entryTypes.append(et)
        ci = common_xls_formats.ColumnInfo(worksheet, hdrs[i], i)

    # Write the headers
    for ec in entryTypes:
        columnInfo = common_xls_formats.ColumnInfo(worksheet, ec.Header, myColumn, 1, 1)
        columnInfo.columnWrite(
            myRow, myColumn, ec.Header, "text", formats.headerFormat(), True
        )
        columns.append(columnInfo)
        myColumn = myColumn + 1

    myColumn = startColumn  # + 1
    myRow = myRow + 1
    chartStartRow = myRow
    ci = columns[0]
    hasValue = False
    # History data is stored from the current to the oldest.
    # Let's print the
    for i in range(len(histData.unitQuantity) - 1, -1, -1):

        if hasValue == False and histData.Quantity(i) <= 0:
            continue

        if hasValue == False:
            hasValue = True
            startIndex = i

        rowDateT = transaction.monthdelta(datetime.date.today(), -i)

        # print(colDateT)
        if (
            calendar.month_abbr[rowDateT.month] == "Jan" or i == startIndex
        ):  # len(histData.unitQuantity)-1:
            rowHeader = calendar.month_abbr[rowDateT.month] + "/" + str(rowDateT.year)
        else:
            rowHeader = calendar.month_abbr[rowDateT.month]

        ci.columnWrite(
            myRow, myColumn, rowHeader, "text", formats.textFormat(myRow), True
        )
        # print(">{}:({}.{}):{}".format(i,myRow,myColumn,rowHeader))
        myRow = myRow + 1

    chartEndRow = myRow

    myRow = startRow + 1

    myColumn = startColumn

    for i in range(startIndex, -1, -1):

        for c in range(1, len(columns)):
            # print("{}:{}".format(i,c))
            ci = columns[c]
            if hdrs[c] == "Price":
                fmt = formats.currencyFormat(myRow)
                value = histData.Price(i)
            elif hdrs[c] == "Quantity":
                fmt = formats.numberFormat(myRow)
                value = histData.Quantity(i)
            elif hdrs[c] == "Value":
                value = histData.Value(i)
                fmt = formats.currencyFormat(myRow)
            elif hdrs[c] == "Dividends":
                value = divData[i]
                fmt = formats.currencyFormat(myRow)
            else:
                print("Invalid colunm {}".format(hdrs[c]))
                raise

            ci.columnWrite(myRow, myColumn + c, value, mTypes[c], fmt)

        myRow = myRow + 1

    for c in columns:
        c.columnSetSize(1.4)

    # chartTypes = [ "Price", "Quantity", "Value"]
    colors = ["#4DA6FF", "#88FF4B", "#B30059", "#FF9900", "#FF9911"]

    # print(">{}".format(ticker.name))
    for i in range(1, len(hdrs)):

        if hdrs[i] == "Dividends":
            myChart = workbook.add_chart({"type": "column"})
        else:
            myChart = workbook.add_chart({"type": "line"})

        myChart.set_title({"name": "{} {} Chart".format(ticker.name, hdrs[i])})
        myChart.set_size({"width": 900, "height": 300})

        columnStart = xlsxwriter.utility.xl_col_to_name(
            startColumn
        )  # this is where the months are
        myLegendValues = (
            "='"
            + worksheetName
            + "'!$"
            + columnStart
            + "$"
            + str(chartStartRow + 1)
            + ":$"
            + columnStart
            + "$"
            + str(chartEndRow)
        )
        # print("myLegendValues:{}".format(myLegendValues))

        columnStart = xlsxwriter.utility.xl_col_to_name(startColumn + i)
        myValues = (
            "='"
            + worksheetName
            + "'!$"
            + columnStart
            + "$"
            + str(chartStartRow + 1)
            + ":$"
            + columnStart
            + "$"
            + str(chartEndRow)
        )

        myChart.add_series(
            {
                "name": hdrs[i],
                "values": myValues,
                "categories": myLegendValues,
                "fill": {"color": colors[i]},
            }
        )

        myChart.set_x_axis(
            {
                "date_axis": True,
                "minor_unit": 4,
                "minor_unit_type": "months",
            }
        )

        # Insert the chart into the worksheet.
        r = int((i - 1) % 2)
        c = int((i - 1) / 2)
        colName = xlsxwriter.utility.xl_col_to_name(startColumn + 5 + (c * 7))

        worksheet.insert_chart(colName + str(startRow + 2 + (r * 17)), myChart)

        #
        for ci in columns:
            ci.columnSetSize(1.4)
        # end charts


#  iexDataSheet - create a worksheet based the data returned from IEX API.
#    worksheet -
#    iexData - an list of tags, labels, types, and formats for the data
#    data_type - the data type from the ticker such as 'quote_data' or 'stats_data'.
#    symbols - The list of tickers built by basedatacsv.py
#    formats - The formats only used by the headers.
def iexDataSheet(worksheet, iexData, data_type, symbols, formats):

    qdColumnInfo = []
    myColumn = 0
    for qd in iexData:
        ci = common_xls_formats.ColumnInfo(worksheet, qd.label, myColumn, 1, 1, 40)
        ci.columnWrite(0, myColumn, qd.label, "text", formats.headerFormat(), True)
        qdColumnInfo.append(ci)
        myColumn = myColumn + 1

    myRow = 1
    for key, value in sorted(symbols.items()):
        t = symbols[key]

        if t.numShares() <= 0:
            continue

        myColumn = 0
        dataSet = t.get_data(data_type)
        # print("qdata:",dataSet)
        if dataSet == None:
            continue

        for qd in iexData:
            myTag = qd.tag
            # print("tag:",myTag)
            dataElement = dataSet.get(myTag)
            ci = qdColumnInfo[myColumn]

            if dataElement != None and (
                qd.type == "currency" or qd.type == "percentage" or qd.type == "number"
            ):
                try:
                    fValue = float(dataElement)
                    ci.columnWrite(myRow, myColumn, fValue, qd.type, qd.format(myRow))
                except:
                    ci.columnWrite(
                        myRow, myColumn, dataElement, qd.type, qd.format(myRow)
                    )

            elif dataElement != None and qd.type == "timestamp":
                tsInt = int(dataElement)
                if tsInt > 0:
                    tsDt = datetime.datetime.fromtimestamp(tsInt / 1000.0)
                    # print(ci.name,":",tsDt)
                    ci.columnWrite(myRow, myColumn, tsDt, "timestamp", qd.format(myRow))
                else:
                    print(
                        "Error: Field ", qd.tag, " has no timestime value:", dataElement
                    )
                    ci.columnWrite(
                        myRow, myColumn, dataElement, "text", qd.format(myRow)
                    )
            else:
                ci.columnWrite(myRow, myColumn, dataElement, qd.type, qd.format(myRow))

            myColumn = myColumn + 1

        myRow = myRow + 1

    #
    for ci in qdColumnInfo:
        # print(ci)
        ci.columnSetSize(1.4)


if __name__ == "__main__":

    #
    # prepare the input csv and excel worksheet file names.
    #
    # inFilename = "transactions.csv"
    # outFilename = "stock_analysis.xlsx"
    # lookupFilename = "lookup.csv"
    # portfolioFilename = "portfolio_value.csv"

    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--input", "-i", help="Input CSV File", default="transactions.csv"
    )
    parser.add_argument(
        "--output", "-o", help="Output XLSX File", default="stock_analysis.xlsx"
    )
    parser.add_argument(
        "--lookup",
        "-l",
        help="File containing lookups for translations",
        default="lookup.csv",
    )
    parser.add_argument(
        "--portfolio",
        "-p",
        help="Portfolio Values Lookups",
        default="portfolio_value.csv",
    )
    args = parser.parse_args()

    # create the workbook
    workbook = xlsxwriter.Workbook(args.output)
    formats = common_xls_formats.XLSFormats(workbook)
    cache = stock_cache.StockCache()

    if not cache.isCouchDBUp():
        print("ERROR: CouchDB is not up")
        exit(-1)

    #
    # Load the Portfolio Value CSV file.  This provides the last price when it's not available
    # through iexdata.
    #
    pValues = pv.PortfolioValue(args.portfolio, args.lookup)
    pValues.WriteValuesWorksheet(workbook, formats)
    pValues.WriteLookupWorksheet(workbook, formats)
    if hasattr(pValues, "created"):
        cache.LoadCacheFromJson("portfolio_value", pValues.data, date=pValues.created)

    #
    # read in the transactions and write them to their own worksheet for any ad-hoc analysis.
    #
    T = transaction.Transactions(workbook, formats)
    T.loadTransactions(args.input, pValues.Lookups())
    T.writeTransactions(0, 0)
    T.getDividends(48)

    translist = T.transactions

    symbols = dict()
    unique_accounts = []

    for row in translist:
        # add row to the master list
        #
        # if row.get_value("symbol") != "HD":
        #    continue
        #
        # if row.get_value("account") != "Ameritrade IRA":
        #    continue

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

        bdata.ProcessEntry(e, symbols, unique_accounts)

    unique_accounts.sort()

    # details are the rows of the Stock Analysis worksheet.
    details = []
    bdata.createSheet(symbols, unique_accounts, details)

    worksheet = workbook.add_worksheet("Stock Analysis")

    # First 2 columns are Name and Symbol
    ColumnInfo = []
    # Name Column
    ci = common_xls_formats.ColumnInfo(worksheet, "Name", 0, 1, 40)
    ci.columnWrite(0, 0, "Name", "text", formats.headerFormat())
    ColumnInfo.append(ci)

    # Symbol Column
    ci = common_xls_formats.ColumnInfo(worksheet, "Symbol", 1, 1, 10)
    ci.columnWrite(0, 1, "Symbol", "text", formats.headerFormat())
    ColumnInfo.append(ci)

    myColumn = 2
    # Add in the individual accounts next
    #
    for acct in unique_accounts:
        # myColumn = xlsxwriter.utility.xl_col_to_name(myColumn) + "1"
        # print("myColumn:",myColumn)
        ci = common_xls_formats.ColumnInfo(worksheet, acct, myColumn, 1, 10)
        ci.columnWrite(0, myColumn, acct, "text", formats.headerFormat(), True)

        # worksheet.write(0,myColumn,acct,formats.headerFormat())
        ColumnInfo.append(ci)
        myColumn = myColumn + 1

    # add in Total Shares
    ci = common_xls_formats.ColumnInfo(worksheet, "Total Shares", myColumn)
    ColumnInfo.append(ci)
    ci.columnWrite(0, myColumn, "Total Shares", "text", formats.headerFormat(), True)
    # worksheet.write(0,myColumn, 'Total Shares',formats.headerFormat())
    totalSharesCol = xlsxwriter.utility.xl_col_to_name(myColumn)
    myColumn = myColumn + 1

    # go through ALL the details and pick up the remaining labels.
    for d in details:
        #  for each detail
        for k in d:
            # for each key in the detail
            dup = False
            # TODO: Update for a better way of making sure there are no duplicates.
            for ci in ColumnInfo:
                if ci.name == k:
                    # print("Skipping ",k," duplicate")
                    dup = True
                    break
                    # for ci

            if dup == True:
                continue

            ci = common_xls_formats.ColumnInfo(worksheet, k, myColumn, 1, 9, 9)
            ColumnInfo.append(ci)
            ci.columnWrite(0, myColumn, k, "text", formats.headerFormat(), True)

            if k == "Latest Price":
                latestPriceCol = xlsxwriter.utility.xl_col_to_name(myColumn)
            elif k == "Total Value":
                totalValueCol = xlsxwriter.utility.xl_col_to_name(myColumn)
            elif k == "Total Cost":
                totalCostCol = xlsxwriter.utility.xl_col_to_name(myColumn)
            elif k == "Yearly Dividend":
                yearlyDividendCol = xlsxwriter.utility.xl_col_to_name(myColumn)
            elif k == "Days Owned":
                daysOwnedCol = xlsxwriter.utility.xl_col_to_name(myColumn)
            elif k == "Dividends Received":
                dividendsRecCol = xlsxwriter.utility.xl_col_to_name(myColumn)
            elif k == "Current Dividend":
                currentDividendCol = xlsxwriter.utility.xl_col_to_name(myColumn)

            myColumn = myColumn + 1

    # push the fields that have a price to the worksheet
    myRow = 1
    for d in details:
        myColumn = 0

        for ci in ColumnInfo:
            # for l in myKeys:
            # if xRow == 2:
            #    print("l:",l)

            if ci.name == "Name":
                s = d["Symbol"]
                ticker = symbols[s]
                myUrl = ticker.worksheetURL()
            else:
                myUrl = None

            v = d[ci.name]
            # myColumn = xlsxwriter.utility.xl_col_to_name(myColumn)
            # currentA = myColumn + str(xRow)

            if myUrl != None:
                #
                # If the URL is present it can only be the Name column.  If the URL wasn't
                # available for the Name, the logic would have it write the Name in the section
                # below
                #
                ci.columnWrite(
                    myRow, myColumn, v, "url", formats.textFormat(myRow), False, myUrl
                )

            else:
                if (
                    ci.name == "Dividends Received"
                    or ci.name == "Total Cost"
                    or ci.name == "Average Price"
                    or ci.name == "Current Dividend"
                    or ci.name == "Yearly Dividend"
                    or ci.name == "Latest EPS"
                ):

                    ci.columnWrite(
                        myRow, myColumn, v, "currency", formats.currencyFormat(myRow)
                    )

                elif ci.name == "Latest Price":
                    #
                    # for mutual funds and international stocks, use the value that is in the portfolio.
                    if d["Latest Price"] != 0:
                        lPrice = d["Latest Price"]
                    else:
                        lsymbol = d.get("Symbol")
                        #
                        lPrice = pValues.GetValue(lsymbol, "quote")
                        #
                        if lPrice == None:
                            lPrice = 0.00
                        # print("Latest Price:",lPrice)

                    ci.columnWrite(
                        myRow,
                        myColumn,
                        lPrice,
                        "currency",
                        formats.currencyFormat(myRow),
                    )

                elif ci.name == "Total Value":
                    # let excel do the work:
                    formula = (
                        "="
                        + str(totalSharesCol)
                        + str(myRow + 1)
                        + "*"
                        + str(latestPriceCol)
                        + str(myRow + 1)
                    )
                    ci.columnWrite(
                        myRow,
                        myColumn,
                        formula,
                        "formula",
                        formats.currencyFormat(myRow),
                    )

                elif ci.name == "Net":
                    # =(L2-N2)+M2
                    # TODO:  Look at why Net is not computing correctly in basedatacsv.py
                    formula = (
                        "=("
                        + str(totalValueCol)
                        + str(myRow + 1)
                        + "-"
                        + str(totalCostCol)
                        + str(myRow + 1)
                        + ") + "
                        + str(dividendsRecCol)
                        + str(myRow + 1)
                    )
                    ci.columnWrite(
                        myRow,
                        myColumn,
                        formula,
                        "formula",
                        formats.currencyFormat(myRow),
                    )

                elif ci.name == "First Purchase":
                    ci.columnWrite(
                        myRow, myColumn, v, "date", formats.dateFormat(myRow)
                    )

                else:
                    ci.columnWrite(
                        myRow, myColumn, v, "text", formats.textFormat(myRow)
                    )

            myColumn = myColumn + 1

        myRow = myRow + 1

    #
    # Add the formulas
    #
    # Name Column
    ci = common_xls_formats.ColumnInfo(
        worksheet, "Projected Dividends", myColumn, 1, 10, 10
    )
    ci.columnWrite(
        0, myColumn, "Projected Dividends", "text", formats.headerFormat(), True
    )
    projectedDividendsCol = xlsxwriter.utility.xl_col_to_name(myColumn)
    ColumnInfo.append(ci)

    ci = common_xls_formats.ColumnInfo(
        worksheet, "Dividend Yield", myColumn + 1, 1, 10, 10
    )
    ci.columnWrite(
        0, myColumn + 1, "Dividend Yield", "text", formats.headerFormat(), True
    )
    ColumnInfo.append(ci)

    ci = common_xls_formats.ColumnInfo(
        worksheet, "Percentage Portfolio", myColumn + 2, 1, 10, 10
    )
    ci.columnWrite(
        0, myColumn + 2, "Percentage Portfolio", "text", formats.headerFormat(), True
    )
    ColumnInfo.append(ci)

    ci = common_xls_formats.ColumnInfo(worksheet, "ROI", myColumn + 3, 1, 10, 10)
    ci.columnWrite(0, myColumn + 3, "ROI", "text", formats.headerFormat())
    ColumnInfo.append(ci)

    ci = common_xls_formats.ColumnInfo(
        worksheet, "Annual Return", myColumn + 4, 1, 10, 10
    )
    ci.columnWrite(
        0, myColumn + 4, "Annual Retrun", "text", formats.headerFormat(), True
    )
    ColumnInfo.append(ci)

    ci = common_xls_formats.ColumnInfo(worksheet, "CAGR", myColumn + 5, 1, 10, 10)
    ci.columnWrite(0, myColumn + 5, "CAGR", "text", formats.headerFormat(), True)
    ColumnInfo.append(ci)

    # special format for the conditial format below
    # TODO: Add these to the common formats
    formatYellow = workbook.add_format({"bg_color": "#FFEB9C", "font_color": "#9C6500"})

    # get the specific columns
    percPortCol = xlsxwriter.utility.xl_col_to_name(myColumn + 2)
    roiCol = xlsxwriter.utility.xl_col_to_name(myColumn + 3)
    aRetCol = xlsxwriter.utility.xl_col_to_name(myColumn + 4)
    cagrCol = xlsxwriter.utility.xl_col_to_name(myColumn + 5)

    worksheet.conditional_format(
        percPortCol + str(2) + ":" + percPortCol + str(myRow),
        {"type": "top", "value": 10, "criteria": "%", "format": formatYellow},
    )

    worksheet.conditional_format(
        roiCol + str(2) + ":" + roiCol + str(myRow),
        {
            "type": "data_bar",
            "bar_no_border": True,
            "bar_color": "#63C384",
            "bar_axis_color": "#0070C0",
        },
    )

    worksheet.conditional_format(
        aRetCol + str(2) + ":" + aRetCol + str(myRow),
        {
            "type": "data_bar",
            "bar_no_border": True,
            "bar_color": "#63C384",
            "bar_axis_color": "#0070C0",
        },
    )

    worksheet.conditional_format(
        cagrCol + str(2) + ":" + cagrCol + str(myRow),
        {
            "type": "data_bar",
            "bar_no_border": True,
            "bar_color": "#63C384",
            "bar_axis_color": "#0070C0",
        },
    )

    # myColumn = stockOrd
    for i in range(2, myRow + 1):

        # Projected Dividends
        # =IF(P17>0,Q17*J17,(M17/U17)*365)
        #  if Yearly Dividends > 0, then Yearly Dividend * Number of shares, else (Total Dividends / Number of Days owned) * 365
        formula = (
            "=if("
            + yearlyDividendCol
            + str(i)
            + "> 0,"
            + yearlyDividendCol
            + str(i)
            + "*"
            + totalSharesCol
            + str(i)
            + ", ("
            + dividendsRecCol
            + str(i)
            + "/"
            + daysOwnedCol
            + str(i)
            + ") * 365)"
        )
        ci = ColumnInfo[myColumn]
        ci.columnWrite(
            i - 1, myColumn, formula, "formula", formats.currencyFormat(i - 1)
        )

        #
        # Dividend Yield:
        #  TODO - Calculate based on dividends recieved in last 12 months.
        # Updated to calculate the yield based on dividends recieved since the mutual funds do not report on dividends.  This
        # is just an approximation for comparison.
        #   Q - Yearly Dividend (yearlyDividendCol)
        #   K - Latest Price (latestPriceCol)
        #   M - Dividends Recieved (dividendsRecCol)
        #   P - Current Dividend (currentDividendCol)
        #   V - Projected Dividends (projectedDividendsCol)
        #   J - Total Shares (totalSharesCol)
        #   K - Latest Price (latestPriceCol)
        # =IF(Q43 > 0, Q43 / K43, IF(M43 > 0, IF(P43=0, (V43 / J43) / K43, 0), 0))
        #  The log is:
        #  If there is a yearly dividend, then divide it by the share price to get the yield
        #  Else:
        #    if there have been dividends and there is no current dividend listed (such as in mutual funds),
        #        then divide the projected divideds by the total shares to get an idea of the dividend per share
        #         and divided it againsts the current share price
        #
        #
        formula = (
            "=if("
            + yearlyDividendCol
            + str(i)
            + "> 0 ,"
            + yearlyDividendCol
            + str(i)
            + "/"
            + latestPriceCol
            + str(i)
            + ", if( "
            + dividendsRecCol
            + str(i)
            + " > 0 , if("
            + currentDividendCol
            + str(i)
            + " = 0, ("
            + projectedDividendsCol
            + str(i)
            + "/"
            + totalSharesCol
            + str(i)
            + ") / "
            + latestPriceCol
            + str(i)
            + ",0),0))"
        )

        ci.columnWrite(
            i - 1, myColumn + 1, formula, "formula", formats.percentFormat(i - 1)
        )

        # Add formula for Percentage of the Portfolio
        formula = (
            "="
            + totalValueCol
            + str(i)
            + "/ (SUM($"
            + totalValueCol
            + "$2:$"
            + totalValueCol
            + "$"
            + str(myRow)
            + "))"
        )
        ci = ColumnInfo[myColumn + 2]
        ci.columnWrite(
            i - 1, myColumn + 2, formula, "formula", formats.percentFormat(i - 1)
        )

        # ROI - Return on Investment
        # IF ( Cost > 0 , totalvalue - total cost / total cost ) / 0
        formula = (
            "=if("
            + totalCostCol
            + str(i)
            + ">0,("
            + totalValueCol
            + str(i)
            + "-"
            + totalCostCol
            + str(i)
            + ") / "
            + totalCostCol
            + str(i)
            + ", 0)"
        )
        ci = ColumnInfo[myColumn + 3]
        ci.columnWrite(
            i - 1, myColumn + 3, formula, "formula", formats.percentFormat(i - 1)
        )

        # Annual Return
        # =if(n2>0,POWER((L2/N2),(365/U2))-1,0)
        formula = (
            "=if("
            + totalCostCol
            + str(i)
            + ">0,power(("
            + totalValueCol
            + str(i)
            + "/"
            + totalCostCol
            + str(i)
            + "),(365 / "
            + daysOwnedCol
            + str(i)
            + "))-1, 0)"
        )
        ci = ColumnInfo[myColumn + 4]
        ci.columnWrite(
            i - 1, myColumn + 4, formula, "formula", formats.percentFormat(i - 1)
        )

        # CAGR - Compound annual growth rate
        # =IF(N37>0,POWER(((L37+M37)/N37),(365/U37))-1,0)
        formula = (
            "=if("
            + totalCostCol
            + str(i)
            + ">0,power((("
            + totalValueCol
            + str(i)
            + "+"
            + dividendsRecCol
            + str(i)
            + ")/"
            + totalCostCol
            + str(i)
            + "),(365 / "
            + daysOwnedCol
            + str(i)
            + "))-1,0)"
        )
        ci = ColumnInfo[myColumn + 5]
        ci.columnWrite(
            i - 1, myColumn + 5, formula, "formula", formats.percentFormat(i - 1)
        )

    #
    for ci in ColumnInfo:
        # print(ci)
        ci.columnSetSize(1.4)

    # Check point...
    print(myRow, ",", myColumn)

    # Load in the Labels for the different tags for the IEX data.
    iexFormats = common_xls_formats.InitType(formats)
    iexQuoteData = common_xls_formats.loadDataLabels(
        "quote_data_labels.csv", iexFormats
    )
    iexStatsData = common_xls_formats.loadDataLabels(
        "stats_data_labels.csv", iexFormats
    )
    iexDividendData = common_xls_formats.loadDataLabels(
        "dividend_data_labels.csv", iexFormats
    )
    iexNewsData = common_xls_formats.loadDataLabels("news_data_labels.csv", iexFormats)
    iexEntryData = common_xls_formats.loadDataLabels(
        "entry_data_labels.csv", iexFormats
    )

    # create a consolidated worksheet of quote data
    worksheet = workbook.add_worksheet("Quote Data")
    iexDataSheet(worksheet, iexQuoteData, "quote_data", symbols, formats)

    worksheet = workbook.add_worksheet("Stats Data")
    iexDataSheet(worksheet, iexStatsData, "stats_data", symbols, formats)

    # Collect the Stock History
    # set up the matrix
    hMatrix = stock_history.HistoryMatrix(workbook, formats, cache, 36, "months")
    hMatrix.CreateHistoryMatrix(T.transactions)
    # print("History Matrix:")
    # print(str(hMatrix))
    hMatrix.writeMatrixWorksheet("Quantity")
    hMatrix.writeMatrixWorksheet("Price")
    hMatrix.writeMatrixWorksheet("Value")

    #
    # Add the individual stock sheets
    #
    maxRow = myRow = 1
    # picker = {}
    for key, value in sorted(symbols.items()):

        t = symbols[key]

        if t.numShares() <= 0:
            continue

        # print('Symbol: ' + key)
        worksheetName = t.worksheetName()
        if len(worksheetName) > 30:
            worksheetName = worksheetName[:30]

        # print("Net Cost: {} ", t.netCost())

        worksheet = workbook.add_worksheet(worksheetName)
        myUrl = "internal: 'Stock Analysis'!A1"
        worksheet.write_url(0, 0, myUrl, None, "Back To Stock Analysis", None)

        try:
            historyRow = hMatrix.getHistoryRow(key)
        except Exception as e:
            print(e)
            continue

        if historyRow != None:
            #
            # print("key:{}:{}".format(key,T.getTickerDividends(key)))
            #
            dividendData = T.getTickerDividends(key)
            if isinstance(dividendData, dict):
                dividendHistory = dividendData.get("dividends")
            else:
                print("{} Missing dividend data".format(key))
                dividendHistory = numpy.zeros((len(historyRow.unitQuantity) + 1))

            printMatrix(
                workbook,
                worksheetName,
                worksheet,
                formats,
                t,
                historyRow,
                dividendHistory,
                0,
                1,
            )
        else:
            print("Error")
            break

        entries = t.entryValues()
        printEntries(worksheet, entries, formats, 0, 6)
        # t.printAccounts()

    # last things, write out the looks-ups.
    # WriteLookupWorkSheet(lookUps, workbook, formats)
    workbook.close()
    print("All done")
    sys.exit(0)
    # all done
