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

    worksheet.write(myRow,myColumn,"Lookup Key",formats.headerFormat())
    worksheet.write(myRow,myColumn+1,"Lookup Value",formats.headerFormat())
    myRow = 1

    myKeys = lookups.keys()
    for k in myKeys:
        v = lookups[k]
        worksheet.write(myRow, myColumn, k,formats.textFormat(myRow))
        worksheet.write(myRow, myColumn+1, v,formats.textFormat(myRow))
        myRow = myRow + 1

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
    myKeys = []
    # myColumn = xlsxwriter.utility.xl_col_to_name(0) + "1" # xl_col_to_name
    worksheet.write(0,0,'Name',formats.headerFormat())
    myKeys.append('Name')

    # myColumn = xlsxwriter.utility.xl_col_to_name(1) + "1"
    worksheet.write(0,1,'Symbol',formats.headerFormat())
    myKeys.append('Symbol')

    myColumn = 2
    # Add in the individual accounts next
    # print(len(unique_accounts))
    for acct in unique_accounts:
        # myColumn = xlsxwriter.utility.xl_col_to_name(myColumn) + "1"
        # print("myColumn:",myColumn)
        worksheet.write(0,myColumn,acct,formats.headerFormat())
        myKeys.append(acct)
        myColumn = myColumn + 1

    # add in Total Shares
    myKeys.append('Total Shares')
    worksheet.write(0,myColumn, 'Total Shares',formats.headerFormat())
    totalSharesCol = xlsxwriter.utility.xl_col_to_name(myColumn)
    myColumn = myColumn + 1

    # go through ALL the details and pick up the remaining labels.
    for d in details:
        # print("d:",d)
        for k in d:
            # print("k:",k)
            try:
                # if it's already in my list of keys
                # continue on.
                myKeys.index(k)
                continue
            except ValueError:
                myKeys.append(k)
                # myColumn = xlsxwriter.utility.xl_col_to_name(myColumn)
                # currentA = myColumn + "1"
                worksheet.write(0,myColumn, k, formats.headerFormat())

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
                # print("adding ",k)

    # push it out to the worksheet
    myRow = 1
    for d in details:
        myColumn = 0

        if d['Latest Price'] == 0:
            # print("Skipping ",d["Name"])
            continue

        for l in myKeys:
            # if xRow == 2:
            #    print("l:",l)

            if l == "Name":
                s = d['Symbol']
                ticker = symbols[s]
                myUrl = ticker.worksheetURL()
            else:
                myUrl = None

            v = d[l]
            # myColumn = xlsxwriter.utility.xl_col_to_name(myColumn)
            # currentA = myColumn + str(xRow)

            if myUrl != None:
                worksheet.write_url(myRow,myColumn,myUrl,formats.textFormat(myRow),v,None)
                # worksheet.write_url(myRow,myColumn,myUrl,None,v,None)
            else:
                if l == "Total Value" or l == "Dividends Received" or l == "Total Cost" or l == "Average Price" or l == "Current Dividend" or l == "Yearly Dividend" or l == "Net" or l == 'Latest EPS':
                    worksheet.write(myRow,myColumn,v,formats.currencyFormat(myRow))
                elif l == 'First Purchase':
                    worksheet.write(myRow, myColumn, v, formats.dateFormat(myRow))
                else:
                    worksheet.write(myRow,myColumn, v, formats.textFormat(myRow))

            myColumn = myColumn + 1

        myRow = myRow + 1

    stockOrd = myColumn
    #  This section writes out the stocks and funds that did not get a price
    #  from the lookup.
    #
    for d in details:
        myColumn = 0

        if d['Latest Price'] != 0:
            # print("Skipping",d['Name'])
            continue

        for l in myKeys:
            v = d[l]
            # myColumn = xlsxwriter.utility.xl_col_to_name(myColumn)
            # currentA = myColumn + str(xRow)

            if l == "Name":
                s = d['Symbol']
                ticker = symbols[s]
                myUrl = ticker.worksheetURL()
            else:
                myUrl = None

            if l == 'Total Value':
                formula = "=" + str(totalSharesCol) + str(myRow+1) + "*" + str(latestPriceCol) + str(myRow+1)
                # print(formula)
                worksheet.write_formula(myRow,myColumn,formula, formats.currencyFormat(myRow))
            elif l == 'Net':
                formula = "=" + totalValueCol + str(myRow) + "-" + totalCostCol + str(myRow)
                worksheet.write_formula(myRow,myColumn,formula, formats.currencyFormat(myRow))
            elif l == 'Latest Price':
                #
                lsymbol = d.get('Symbol')
                #
                lpriceV = pValues.get(lsymbol)
                #
                if lpriceV != None:
                    lprice = lpriceV.get('quote')
                else:
                    lprice = 0.00
                worksheet.write_number(myRow,myColumn, float(lprice),formats.currencyFormat(myRow))
            elif (l == 'Current Dividend' or l == 'Yearly Dividend' or l == 'Latest EPS' or l == 'Dividends Received' or l == 'Total Cost' or l == 'Average Price') and v != None:
                worksheet.write_number(myRow, myColumn, float(v), formats.currencyFormat(myRow))
            elif l == 'First Purchase':
                worksheet.write(myRow, myColumn, v, formats.dateFormat(myRow))
            else:
                if myUrl != None:
                    worksheet.write_url(myRow, myColumn, myUrl, formats.textFormat(myRow), v, None)
                else:
                    worksheet.write(myRow, myColumn, v, formats.textFormat(myRow))
                #

            myColumn = myColumn + 1

        myRow = myRow + 1

    #
    #  If the last section didn't have any entries, myColumn will be 0
    #  reset it to the last value of the previous section.
    #
    if myColumn == 0:
        myColumn = stockOrd

    worksheet.write(0,myColumn,'Percentage Portfolio',formats.headerFormat())
    worksheet.write(0,myColumn+1,'Dividend Yield',formats.headerFormat())
    worksheet.write(0,myColumn+2,'ROI',formats.headerFormat())
    worksheet.write(0,myColumn+3,'Annual Return',formats.headerFormat())
    worksheet.write(0,myColumn+4,'CAGR', formats.headerFormat())
    worksheet.write(0,myColumn+5,'Projected Dividends',formats.headerFormat())

    # myColumn = stockOrd
    for i in range(2,myRow+1):

        # Add formula for Percentage of the Portfolio
        currentB =  totalValueCol + str(i)
        formula = '=' + currentB + '/ (SUM($' + totalValueCol + "$2:$" + totalValueCol + "$" + str(myRow) + "))"
        worksheet.write_formula(i-1,myColumn,formula,formats.percentFormat(i-1))

        # Dividend Yield
        # IF(Q2 > 0 , Q2/K2, 0 )
        formula = "=IF(" + yearlyDividendCol + str(i) + "> 0 ," + yearlyDividendCol + str(i) + "/" + latestPriceCol + str(i) + ",0)"
        worksheet.write_formula(i-1,myColumn+1,formula,formats.percentFormat(i-1))

        # ROI - Return on Investment
        # IF ( Cost > 0 , totalvalue - total cost / total cost ) / 0
        formula = "=if(" + totalCostCol + str(i) + ">0,(" + totalValueCol + str(i) + "-" + totalCostCol + str(i) + ") / " + totalCostCol + str(i) + ", 0)"
        worksheet.write_formula(i-1,myColumn+2,formula,formats.percentFormat(i-1))

        # Annual Return
        # =if(n2>0,POWER((L2/N2),(365/U2))-1,0)
        formula = "=if(" + totalCostCol + str(i) + ">0,power((" + totalValueCol + str(i) + "/" + totalCostCol + str(i) + "),(365 / " + daysOwnedCol + str(i) + "))-1, 0)"
        worksheet.write_formula(i-1,myColumn+3,formula,formats.percentFormat(i-1))

        # CAGR - Compound annual growth rate
        # =IF(N37>0,POWER(((L37+M37)/N37),(365/U37))-1,0)
        formula = "=if(" + totalCostCol + str(i) + ">0,power(((" + totalValueCol + str(i) + "+" + dividendsRecCol + str(i) + ")/" + totalCostCol + str(i) + "),(365 / " + daysOwnedCol + str(i) + "))-1,0)"
        worksheet.write_formula(i-1,myColumn+4,formula,formats.percentFormat(i-1))
        #
        # Projected Dividends
        # =IF(P17>0,Q17*J17,(M17/U17)*365)
        #  if Yearly Dividends > 0, then Yearly Dividend * Number of shares, else (Total Dividends / Number of Days owned) * 365
        formula = "=if(" + yearlyDividendCol + str(i) + "> 0," + yearlyDividendCol + str(i) + "*" + totalSharesCol + str(i) + ", (" + dividendsRecCol + str(i) + "/" + daysOwnedCol + str(i) + ") * 365)"
        # print(formula)
        worksheet.write_formula(i-1, myColumn+5, formula, formats.currencyFormat(i-1))

    # marker
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
