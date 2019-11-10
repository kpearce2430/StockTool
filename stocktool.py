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
import basedatacsv as bdata
import portfoliovalue as pvalue
import transaction
import urllib3
import datetime
import common_xls_formats

def LoadLookup(name, lookup):
    lookupReader = csv.reader(open(name,newline=''),delimiter=',', quotechar='"')
    for row in lookupReader:
        if len(row) == 2:
            # print(row[0],row[1])
            lookup[row[0]] = row[1]
        else:
            print("huh:",row)

def WriteLookupWorkSheet( lookups, workbook, formats, startRow = 0, startCol = 0):

    myRow = startRow
    myColumn = startCol

    worksheet = workbook.add_worksheet('Lookups')

    keyName = "Lookup Key"
    ciKey = common_xls_formats.ColumnInfo(worksheet,keyName,myColumn,1,1,100)
    ciKey.columnWrite(myRow,myColumn,keyName,'text',formats.headerFormat())

    valueName = "Lookup Value"
    ciValue = common_xls_formats.ColumnInfo(worksheet,valueName,myColumn+1,1,1,100)
    ciValue.columnWrite(myRow,myColumn+1,valueName,'text',formats.headerFormat(),True)

    myRow = myRow + 1
    myKeys = lookups.keys()
    for k in myKeys:
        v = lookups[k]
        ciKey.columnWrite(myRow,myColumn,k,'text',formats.textFormat(myRow))
        ciValue.columnWrite(myRow,myColumn+1,v,'text',formats.textFormat(myRow))
        myRow = myRow + 1

    ciKey.columnSetSize(1.4)
    ciValue.columnSetSize(1.4)

def printData( labelData, ticker, formats, row = 0, col = 0 ):
    myRow = row;
    myColumn = col

    for d in labelData:
        tag = d.tag
        field = ticker.get(tag)
        if field != None:
            worksheet.write(myRow, myColumn, d.label, formats.textFormat(myRow))
            if d.format != None:
                # print(d.tag,",",d.format)
                if d.type == 'timestamp':
                    tsInt = int(field)
                    if tsInt > 0:
                        tsDt = datetime.datetime.fromtimestamp(tsInt / 1000.0)
                        worksheet.write(myRow, myColumn + 1, tsDt, d.format(myRow))
                    else:
                        print("Field ", d.tag, " has no timestime value:", field)
                        worksheet.write(myRow, myColumn + 1, field, formats.textFormat(myRow))
                else:
                    worksheet.write(myRow, myColumn + 1, field, d.format(myRow))
            else:
                print("Missing format for ", d.tag, ",", d.type)
                worksheet.write(myRow, myColumn + 1, field, None )
            myRow = myRow + 1

    return myRow

if __name__ == "__main__":

    #
    # prepare the input csv and excel worksheet file names.
    #
    inFilename = "transactions.csv"
    outFilename = "stock_analysis.xlsx"
    lookupFilename = "lookup.csv"
    portfolioFilename = "portfolio_value.csv"

    i = 0
    for i in range(1, len(sys.argv)):

        if i == 1:
            inFilename = sys.argv[i]
        elif i == 2:
            outFilename = sys.argv[i]
        elif i == 3:
            lookupFilename = sys.argv[i]
        elif i == 4:
            portfolioFilename = sys.argv[i]
        else:
            print("Ignoring extra arguments", sys.argv[i])

    # create the workbook
    workbook = xlsxwriter.Workbook(outFilename)
    formats = common_xls_formats.XLSFormats(workbook)

    # load up the lookup table
    lookUps = dict()
    LoadLookup(lookupFilename, lookUps)

    # load the portfolio value
    #
    # Load the Portfolio Value CSV file.  This provides the last price when it's not available
    # through iexdata.
    pValues = dict()
    pvalue.LoadPortfolioValue(portfolioFilename,pValues, lookUps)
    pvalue.WritePortfolioValueWorksheet(pValues,workbook,formats)

    #
    # read in the transactions and write them to their own worksheet for any ad-hoc analysis.
    #
    T = transaction.Transactions(workbook,formats)
    T.loadTransactions(inFilename,lookUps)
    T.writeTransactions(0,0)
    translist = T.transactions

    #
    # set up the http client to pull stock data.
    #
    http = urllib3.PoolManager()
    urllib3.disable_warnings()

    symbols = dict()
    unique_accounts = []

    for row in translist:
        # add row to the master list
        #
        # if row.get_value("symbol") != "HD":
        #   continue

        amt = row.get_value("amount")
        if row.get_value("type") == 'Reinvest Dividend':
            amt = row.get_value("invest_amt")

        e = bdata.Entry(
            row.get_value("date"),
            row.get_value("type"),
            row.get_value("security_payee"),
            row.get_value("security"),
            row.get_value("description"),
            row.get_value("shares"),
            amt,
            row.get_value("account"))

        bdata.ProcessEntry(e,symbols,unique_accounts,http)

    unique_accounts.sort()

    # details are the rows of the Stock Analysis worksheet.
    details = []
    bdata.createSheet(symbols, unique_accounts, details)

    worksheet = workbook.add_worksheet('Stock Analysis')

    # First 2 columns are Name and Symbol
    ColumnInfo = []
    # Name Column
    ci = common_xls_formats.ColumnInfo(worksheet,"Name",0,1,40)
    ci.columnWrite(0,0,"Name",'text',formats.headerFormat())
    ColumnInfo.append(ci)

    # Symbol Column
    ci = common_xls_formats.ColumnInfo(worksheet,"Symbol",1,1,10)
    ci.columnWrite(0,1,"Symbol",'text',formats.headerFormat())
    ColumnInfo.append(ci)

    myColumn = 2
    # Add in the individual accounts next
    # print(len(unique_accounts))
    for acct in unique_accounts:
        # myColumn = xlsxwriter.utility.xl_col_to_name(myColumn) + "1"
        # print("myColumn:",myColumn)
        ci = common_xls_formats.ColumnInfo(worksheet,acct,myColumn,1,10)
        ci.columnWrite(0,myColumn,acct,'text',formats.headerFormat(),True)

        # worksheet.write(0,myColumn,acct,formats.headerFormat())
        ColumnInfo.append(ci)
        myColumn = myColumn + 1

    # add in Total Shares
    ci = common_xls_formats.ColumnInfo(worksheet,'Total Shares',myColumn)
    ColumnInfo.append(ci)
    ci.columnWrite(0,myColumn,'Total Shares','text',formats.headerFormat(),True)
    # worksheet.write(0,myColumn, 'Total Shares',formats.headerFormat())
    totalSharesCol = xlsxwriter.utility.xl_col_to_name(myColumn)
    myColumn = myColumn + 1

    # go through ALL the details and pick up the remaining labels.
    for d in details:
        # print("d:",d)
        for k in d:
            # print("k:",k)

            dup = False
            for ci in ColumnInfo:
                if ci.name == k:
                    print("Skipping ",k," duplicate")
                    dup = True
                    break; # for ci

            if dup == True:
                continue

            ci = common_xls_formats.ColumnInfo(worksheet,k,myColumn,1,9,9)
            ColumnInfo.append(ci)
            ci.columnWrite(0,myColumn,k,'text',formats.headerFormat(),True)

            if k == 'Latest Price':
                latestPriceCol = xlsxwriter.utility.xl_col_to_name(myColumn)
            elif k == 'Total Value':
                totalValueCol = xlsxwriter.utility.xl_col_to_name(myColumn)
            elif k == 'Total Cost':
                totalCostCol = xlsxwriter.utility.xl_col_to_name(myColumn)
            elif k == 'Yearly Dividend':
                yearlyDividendCol = xlsxwriter.utility.xl_col_to_name(myColumn)
            elif k == 'Days Owned':
                daysOwnedCol = xlsxwriter.utility.xl_col_to_name(myColumn)
            elif k == 'Dividends Received':
                dividendsRecCol = xlsxwriter.utility.xl_col_to_name(myColumn)

            myColumn = myColumn + 1

    # push the fields that have a price to the worksheet
    myRow = 1
    for d in details:
        myColumn = 0

        # if d['Latest Price'] == 0:
        #    print("Skipping ",d["Name"])
        #    continue

        for ci in ColumnInfo:
        # for l in myKeys:
            # if xRow == 2:
            #    print("l:",l)

            if ci.name == "Name":
                s = d['Symbol']
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
                ci.columnWrite(myRow,myColumn,v,'url',formats.textFormat(myRow),False,myUrl)

            else:
                if  ci.name == "Dividends Received" or ci.name == "Total Cost" or ci.name == "Average Price" or \
                        ci.name == "Current Dividend" or ci.name == "Yearly Dividend" or  ci.name == 'Latest EPS':

                    ci.columnWrite(myRow,myColumn,v,'currency',formats.currencyFormat(myRow))

                elif ci.name == 'Latest Price':
                    if d['Latest Price'] != 0:
                        lPrice = d['Latest Price']
                    else:
                        lsymbol = d.get('Symbol')
                        #
                        lpriceV = pValues.get(lsymbol)
                        #
                        if lpriceV != None:
                            lPrice = lpriceV.get('quote')
                        else:
                            lPrice = 0.00

                    ci.columnWrite(myRow,myColumn,lPrice,'currency',formats.currencyFormat(myRow))

                elif ci.name == "Total Value":
                    # let excel do the work:
                    formula = "=" + str(totalSharesCol) + str(myRow + 1) + "*" + str(latestPriceCol) + str(myRow + 1)
                    ci.columnWrite(myRow,myColumn,formula,'formula',formats.currencyFormat(myRow))

                elif ci.name == 'Net':
                    # =(L2-N2)+M2
                    # TODO:  Look at why Net is not computing correctly in basedatacsv.py
                    formula = "=(" + str(totalValueCol) + str(myRow + 1) + "-" + str(totalCostCol) + str(myRow + 1) + ") + " + str(dividendsRecCol) + str(myRow+1)
                    ci.columnWrite(myRow,myColumn,formula,'formula',formats.currencyFormat(myRow))

                elif ci.name == 'First Purchase':
                    ci.columnWrite(myRow,myColumn, v,'date',formats.dateFormat(myRow))

                else:
                    ci.columnWrite(myRow,myColumn, v,'text',formats.textFormat(myRow))

            myColumn = myColumn + 1

        myRow = myRow + 1

    #
    # Add the formulas
    #
    # Name Column
    ci = common_xls_formats.ColumnInfo(worksheet,"Projected Dividends",myColumn,1,10,10)
    ci.columnWrite(0,myColumn,"Projected Dividends",'text',formats.headerFormat(),True)
    ColumnInfo.append(ci)

    ci = common_xls_formats.ColumnInfo(worksheet,"Dividend Yield",myColumn+1,1,10,10)
    ci.columnWrite(0,myColumn+1,"Dividend Yield",'text',formats.headerFormat(),True)
    ColumnInfo.append(ci)

    ci = common_xls_formats.ColumnInfo(worksheet,"Percentage Portfolio",myColumn+2,1,10,10)
    ci.columnWrite(0,myColumn+2,"Percentage Portfolio",'text',formats.headerFormat(),True)
    ColumnInfo.append(ci)

    ci = common_xls_formats.ColumnInfo(worksheet, "ROI", myColumn+3, 1, 10, 10)
    ci.columnWrite(0, myColumn+3, "ROI", 'text', formats.headerFormat())
    ColumnInfo.append(ci)

    ci = common_xls_formats.ColumnInfo(worksheet, "Annual Return", myColumn+4, 1, 10, 10)
    ci.columnWrite(0, myColumn+4, "Annual Retrun", 'text', formats.headerFormat(),True)
    ColumnInfo.append(ci)

    ci = common_xls_formats.ColumnInfo(worksheet, "CAGR", myColumn+5, 1, 10, 10)
    ci.columnWrite(0, myColumn+5, "CAGR", 'text', formats.headerFormat(),True)
    ColumnInfo.append(ci)

    # worksheet.write(0,myColumn,'Percentage Portfolio',formats.headerFormat())
    # worksheet.write(0,myColumn+1,'Dividend Yield',formats.headerFormat())
    # worksheet.write(0,myColumn+2,'ROI',formats.headerFormat())
    # worksheet.write(0,myColumn+3,'Annual Return',formats.headerFormat())
    # worksheet.write(0,myColumn+4,'CAGR', formats.headerFormat())
    # worksheet.write(0,myColumn+5,'Projected Dividends',formats.headerFormat())

    # special format for the conditial format below
    formatYellow = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})

    # get the specific columns
    percPortCol = xlsxwriter.utility.xl_col_to_name(myColumn+2)
    roiCol = xlsxwriter.utility.xl_col_to_name(myColumn+3)
    aRetCol = xlsxwriter.utility.xl_col_to_name(myColumn+4)
    cagrCol = xlsxwriter.utility.xl_col_to_name(myColumn+5)

    worksheet.conditional_format(percPortCol+str(2)+':'+percPortCol+str(myRow), {'type': 'top',
                                           'value': 10,
                                           'criteria': '%',
                                           'format': formatYellow})

    worksheet.conditional_format(roiCol + str(2) + ':' + roiCol + str(myRow), {'type': 'data_bar',
                                                                                 'bar_no_border': True,
                                                                                 'bar_color': '#63C384',
                                                                                 'bar_axis_color': '#0070C0'})

    worksheet.conditional_format(aRetCol+str(2)+':'+aRetCol+str(myRow), {'type': 'data_bar',
                                                                         'bar_no_border': True,
                                                                         'bar_color': '#63C384',
                                                                         'bar_axis_color': '#0070C0'})

    worksheet.conditional_format(cagrCol+str(2)+':'+cagrCol+str(myRow), {'type': 'data_bar',
                                                                         'bar_no_border': True,
                                                                         'bar_color': '#63C384',
                                                                         'bar_axis_color': '#0070C0'})

    # myColumn = stockOrd
    for i in range(2,myRow+1):

        # Projected Dividends
        # =IF(P17>0,Q17*J17,(M17/U17)*365)
        #  if Yearly Dividends > 0, then Yearly Dividend * Number of shares, else (Total Dividends / Number of Days owned) * 365
        formula = "=if(" + yearlyDividendCol + str(i) + "> 0," + yearlyDividendCol + str(i) + "*" + totalSharesCol + str(i) + ", (" + dividendsRecCol + str(i) + "/" + daysOwnedCol + str(i) + ") * 365)"
        ci = ColumnInfo[myColumn]
        ci.columnWrite(i-1,myColumn,formula,'formula',formats.currencyFormat(i-1))

        # Dividend Yield
        # IF(Q2 > 0 , Q2/K2, 0 )
        formula = "=IF(" + yearlyDividendCol + str(i) + "> 0 ," + yearlyDividendCol + str(i) + "/" + latestPriceCol + str(i) + ",0)"
        ci = ColumnInfo[myColumn+1]
        ci.columnWrite(i-1,myColumn+1,formula,'formula',formats.percentFormat(i-1))

        # Add formula for Percentage of the Portfolio
        formula = '=' + totalValueCol + str(i) + '/ (SUM($' + totalValueCol + "$2:$" + totalValueCol + "$" + str(myRow) + "))"
        ci = ColumnInfo[myColumn+2]
        ci.columnWrite(i-1,myColumn+2,formula,'formula',formats.percentFormat(i-1))

        # ROI - Return on Investment
        # IF ( Cost > 0 , totalvalue - total cost / total cost ) / 0
        formula = "=if(" + totalCostCol + str(i) + ">0,(" + totalValueCol + str(i) + "-" + totalCostCol + str(i) + ") / " + totalCostCol + str(i) + ", 0)"
        ci = ColumnInfo[myColumn+3]
        ci.columnWrite(i-1,myColumn+3,formula,'formula',formats.percentFormat(i-1))

        # Annual Return
        # =if(n2>0,POWER((L2/N2),(365/U2))-1,0)
        formula = "=if(" + totalCostCol + str(i) + ">0,power((" + totalValueCol + str(i) + "/" + totalCostCol + str(i) + "),(365 / " + daysOwnedCol + str(i) + "))-1, 0)"
        ci = ColumnInfo[myColumn+4]
        ci.columnWrite(i-1,myColumn+4,formula,'formula',formats.percentFormat(i-1))

        # CAGR - Compound annual growth rate
        # =IF(N37>0,POWER(((L37+M37)/N37),(365/U37))-1,0)
        formula = "=if(" + totalCostCol + str(i) + ">0,power(((" + totalValueCol + str(i) + "+" + dividendsRecCol + str(i) + ")/" + totalCostCol + str(i) + "),(365 / " + daysOwnedCol + str(i) + "))-1,0)"
        ci = ColumnInfo[myColumn+5]
        ci.columnWrite(i-1,myColumn+5,formula,'formula',formats.percentFormat(i-1))


    #
    for ci in ColumnInfo:
        print(ci)
        ci.columnSetSize(1.4)

    print(myRow,",",myColumn)

    iexFormats = common_xls_formats.InitType(formats)
    iexQuoteData = common_xls_formats.loadDataLabels("quote_data_labels.csv", iexFormats)
    iexStatsData = common_xls_formats.loadDataLabels("stats_data_labels.csv", iexFormats)
    iexDividendData = common_xls_formats.loadDataLabels("dividend_data_labels.csv",iexFormats)
    iexNewsData = common_xls_formats.loadDataLabels("news_data_labels.csv",iexFormats)

    maxRow = myRow
    for key, value in sorted(symbols.items()):

        t = symbols[key]

        if t.numShares() <= 0:
            continue

        # print('Symbol: ' + key)
        worksheetName = t.worksheetName()
        if len(worksheetName) > 30 :
            worksheetName = worksheetName[:30]

        worksheet = workbook.add_worksheet(worksheetName)
        myUrl = "internal: 'Stock Analysis'!A1"
        worksheet.write_url(0,0,myUrl,None,"Back To Stock Analysis", None)

        myRow =3
        myColumn = 0
        maxRow = myRow

        for data_type in ['quote_data','stats_data','dividend_data','news_data']:

            qdata = t.get_data(data_type)
            if qdata == None:
                continue
            # print(qdata)
            if data_type == 'quote_data':
                worksheet.write(myRow, myColumn, "Quote Data", formats.headerFormat())
                worksheet.write(myRow, myColumn + 1, "", formats.headerFormat())
                myRow = printData(iexQuoteData,qdata,formats,myRow+1,myColumn)

            elif data_type == 'stats_data':
                worksheet.write(myRow, myColumn, "Stats Data", formats.headerFormat())
                worksheet.write(myRow, myColumn + 1, "", formats.headerFormat())
                myRow = printData(iexStatsData,qdata,formats,myRow+1,myColumn)

            elif data_type == 'dividend_data':
                count = 1
                for div in qdata:
                    worksheet.write(myRow, myColumn, "Dividend Data", formats.headerFormat())
                    worksheet.write(myRow, myColumn + 1, str(count), formats.headerFormat())
                    myRow = printData(iexDividendData,div,formats,myRow+1,myColumn)
                    count = count+1

            elif data_type == 'news_data':
                count = 1
                for div in qdata:
                    worksheet.write(myRow, myColumn, "News Data", formats.headerFormat())
                    worksheet.write(myRow, myColumn + 1, str(count), formats.headerFormat())
                    myRow = printData(iexNewsData,div,formats,myRow+1,myColumn)
                    count = count+1

            elif isinstance(qdata,dict):
                for item in qdata:
                    # print(key,":",qdata[item])
                    worksheet.write(myRow,myColumn,item,formats.textFormat(myRow))
                    worksheet.write(myRow,myColumn+1,qdata[item],formats.textFormat(myRow))
                    myRow = myRow + 1

            else:
                if isinstance(qdata,list) and (data_type == 'dividend_data' or data_type == 'news_data'):
                    count = 1
                    for div in qdata:
                        worksheet.write(myRow,myColumn,data_type ,formats.headerFormat())
                        worksheet.write(myRow,myColumn + 1,str(count),formats.headerFormat())
                        myRow = myRow+1
                        for  datam in div:
                            # print(key,":",qdata[item])
                            worksheet.write(myRow, myColumn, datam, formats.textFormat(myRow))
                            worksheet.write(myRow, myColumn + 1, div[datam], formats.textFormat(myRow))
                            myRow = myRow + 1
                        count = count + 1
                else:
                    print(key,":",data_type,":",qdata)

            if myRow > maxRow:
                maxRow = myRow

            myRow = 3
            myColumn = myColumn + 3

        # print("writing transactions for:",key," maxRow:",maxRow)
        T.writeTransactions(maxRow+1,0,worksheet,transaction.pickSymbol,key)

    # last things, write out the looks-ups.
    WriteLookupWorkSheet(lookUps,workbook,formats)
    workbook.close()
    print("All done")
    sys.exit(0)
    # all done
