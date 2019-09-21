
#
# ----------------------------------------------------------------------------
# "THE BEER-WARE LICENSE" (Revision 42):
# As long as you retain this notice you
# can do whatever you want with this stuff. If we meet some day, and you think
# this stuff is worth it, you can buy me a beer in return.
# ----------------------------------------------------------------------------
#
import sys
import csv
# import json
import xlsxwriter
import basedatacsv as bdata
import portfoliovalue as pvalue
import transaction
import urllib3
import datetime


def LoadLookup(name, lookup):
    lookupReader = csv.reader(open(name,newline=''),delimiter=',', quotechar='"')
    for row in lookupReader:
        if len(row) == 2:
            # print(row[0],row[1])
            lookup[row[0]] = row[1]
        else:
            print("huh:",row)

def WriteLookupWorkSheet( lookups, workbook):
    myRow = 0

    worksheet = workbook.add_worksheet('Lookups')

    myKeys = lookups.keys()
    for k in myKeys:
        v = lookups[k]
        worksheet.write(myRow, 0, k)
        worksheet.write(myRow, 1, v)
        myRow = myRow + 1


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

    # load up the lookup table
    lookUps = dict()
    LoadLookup(lookupFilename, lookUps)
    WriteLookupWorkSheet(lookUps,workbook)

    # load the portfolio value
    #
    # Load the Portfolio Value CSV file.  This provides the last price when it's not available
    # through iexdata.
    pValues = dict()
    pvalue.LoadPortfolioValue(portfolioFilename,pValues, lookUps)
    pvalue.WritePortfolioValueWorksheet(pValues,workbook)

    #
    # read in the transactions and write them to their own worksheet for any ad-hoc analysis.
    #
    translist = []
    transaction.LoadTransactions(inFilename,translist,lookUps)
    transaction.WriteTransactionWorksheet(translist,workbook)

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
        # "date", "type", "security","security_payee","description","shares","amount","account"
        #
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
    myColumn = xlsxwriter.utility.xl_col_to_name(0) + "1" # xl_col_to_name
    worksheet.write(myColumn,'Name')
    myKeys.append('Name')

    myColumn = xlsxwriter.utility.xl_col_to_name(1) + "1"
    worksheet.write(myColumn,'Symbol')
    myKeys.append('Symbol')

    myColumnNum = 2
    # Add in the individual accounts next
    # print(len(unique_accounts))
    for a in unique_accounts:
        myColumn = xlsxwriter.utility.xl_col_to_name(myColumnNum) + "1"
        # print("myColumn:",myColumn)
        worksheet.write(myColumn,a)
        myKeys.append(a)
        myColumnNum = myColumnNum + 1

    # add in Total Shares
    myKeys.append('Total Shares')
    totalSharesCol = xlsxwriter.utility.xl_col_to_name(myColumnNum)
    myColumn = totalSharesCol + "1" # xl_col_to_name
    worksheet.write(myColumn, 'Total Shares')
    myColumnNum = myColumnNum + 1

    # go through ALL the details and pick up the remaining labels.
    for d in details:
        for k in d:
            try:
                # if it's already in my list of keys
                # continue on.
                myKeys.index(k)
                continue
            except ValueError:
                myKeys.append(k)
                myColumn = xlsxwriter.utility.xl_col_to_name(myColumnNum)
                currentA = myColumn + "1"
                worksheet.write(currentA, k)

                if k == 'Latest Price':
                    latestPriceCol = myColumn
                elif k == 'Total Value':
                    totalValueCol = myColumn
                elif k == 'Total Cost':
                    totalCostCol = myColumn
                elif k == 'Yearly Dividend':
                    yearlyDividendCol = myColumn
                elif k == 'Days Owned':
                    daysOwnedCol = myColumn
                elif k == 'Dividends Received':
                    dividendsRecCol = myColumn

                myColumnNum = myColumnNum + 1
                # print("adding ",k)


    # push it out to the worksheet
    xRow = 2
    for d in details:
        myColumnNum = 0

        if d['Latest Price'] == 0:
            # print("Skipping ",d["Name"])
            continue

        for l in myKeys:
            v = d[l]
            myColumn = xlsxwriter.utility.xl_col_to_name(myColumnNum)
            currentA = myColumn + str(xRow)
            worksheet.write(currentA,v)
            myColumnNum = myColumnNum + 1

        xRow = xRow + 1

    stockOrd = myColumnNum

    # worksheet = workbook.add_worksheet('Fund Analysis')

    # xRow = 2
    # myOrd = ord('A')
    # for k in myKeys:
    #    currentA = str(chr(myOrd)) + "1"
    #    worksheet.write(currentA,k)
    #    myOrd = myOrd + 1

    #  This section writes out the stocks and funds that did not get a price
    #  from the lookup.
    #
    for d in details:
        myColumnNum = 0

        # myOrd = ord('A')

        if d['Latest Price'] != 0:
            # print("Skipping",d['Name'])
            continue

        for l in myKeys:
            v = d[l]
            myColumn = xlsxwriter.utility.xl_col_to_name(myColumnNum)
            currentA = myColumn + str(xRow)

            if l == 'Total Value':
                formula = "=" + totalSharesCol + str(xRow) + "*" + latestPriceCol + str(xRow)
                # print(formula)
                worksheet.write_formula(currentA,formula)
            elif l == 'Net':
                formula = "=" + totalValueCol + str(xRow) + "-" + totalCostCol + str(xRow)
                worksheet.write_formula(currentA, formula)
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
                worksheet.write_number(currentA, float(lprice))

            else:
                worksheet.write(currentA, v)

            myColumnNum = myColumnNum + 1

        xRow = xRow + 1


    #
    #  If the last section didn't have any entries, myColumnNum will be 0
    #  reset it to the last value of the previous section.
    if myColumnNum == 0:
        myColumnNum = stockOrd
    #else
    #    stockOrd = myColumnNum

    # add the formulas for
    # ... Percentage of the Portfolio

    percent_fmt = workbook.add_format({'num_format': '0.00%'})

    myColumn = xlsxwriter.utility.xl_col_to_name(myColumnNum)
    currentA = myColumn + "1"
    worksheet.write(currentA,'Percentage Portfolio')

    currentA= myColumn + ":" + myColumn
    worksheet.set_column( currentA, None, percent_fmt)
    # print(str(chr(myOrd)),xRow)

    # myColumnNum = myColumnNum+1
    myColumn = xlsxwriter.utility.xl_col_to_name(myColumnNum+1)
    currentA = myColumn + "1"
    worksheet.write(currentA,'Dividend Yield')
    currentA= myColumn + ":" + myColumn
    worksheet.set_column( currentA, None, percent_fmt)
    # print(str(chr(myOrd+1)),xRow)

    # myColumnNum = myColumnNum + 1
    myColumn = xlsxwriter.utility.xl_col_to_name(myColumnNum+2)
    currentA = myColumn + "1"
    worksheet.write(currentA,'ROI')
    currentA= myColumn + ":" + myColumn
    worksheet.set_column( currentA, None, percent_fmt)
    # print(str(chr(myOrd+2)),xRow)

    # myColumnNum = myColumnNum + 1
    myColumn = xlsxwriter.utility.xl_col_to_name(myColumnNum+3)
    currentA = myColumn + "1"
    worksheet.write(currentA,'Annual Return')
    currentA= myColumn + ":" + myColumn
    worksheet.set_column( currentA, None, percent_fmt)
    # print(str(chr(myOrd+3)),xRow)

    # myColumnNum = myColumnNum + 1
    myColumn = xlsxwriter.utility.xl_col_to_name(myColumnNum+4)
    currentA = myColumn + "1"
    worksheet.write(currentA,'CAGR')
    currentA= myColumn + ":" + myColumn
    worksheet.set_column( currentA, None, percent_fmt)
    print(myColumn,xRow)

    # myColumnNum = stockOrd
    for i in range(2,xRow):

        myColumn = xlsxwriter.utility.xl_col_to_name(myColumnNum)
        currentA = myColumn + str(i)

        currentB =  totalValueCol + str(i)
        formula = '=' + currentB + '/ (SUM($' + totalValueCol + "$2:$" + totalValueCol + "$" + str(xRow-1) + "))"
        # print(formula)
        worksheet.write_formula(currentA,formula)

        #IF(K48>0,R48/K48,0)
        myColumn = xlsxwriter.utility.xl_col_to_name(myColumnNum+1)
        currentA = myColumn + str(i)

        formula = "=if(" + yearlyDividendCol + str(i) + "> 0 ," + yearlyDividendCol + str(i) + "/" + latestPriceCol + str(i) + ",0)"
        worksheet.write_formula(currentA,formula)

        # IF ( Cost > 0 , totalvalue - total cost / total cost ) / 0
        myColumn = xlsxwriter.utility.xl_col_to_name(myColumnNum+2)
        currentA = myColumn + str(i)

        formula = "=if(" + totalCostCol + str(i) + ">0,(" + totalValueCol + str(i) + "-" + totalCostCol + str(i) + ") / " + totalCostCol + str(i) + ", 0)"
        worksheet.write_formula(currentA,formula)

        # =if(n2>0,POWER((L2/N2),(365/U2))-1,0)
        myColumn = xlsxwriter.utility.xl_col_to_name(myColumnNum+3)
        currentA = myColumn + str(i)

        formula = "=if(" + totalCostCol + str(i) + ">0,power((" + totalValueCol + str(i) + "/" + totalCostCol + str(i) + "),(365 / " + daysOwnedCol + str(i) + "))-1, 0)"
        worksheet.write_formula(currentA,formula)

        # =IF(N37>0,POWER(((L37+M37)/N37),(365/U37))-1,0)
        myColumn = xlsxwriter.utility.xl_col_to_name(myColumnNum+4)
        currentA = myColumn + str(i)

        formula = "=if(" + totalCostCol + str(i) + ">0,power(((" + totalValueCol + str(i) + "+" + dividendsRecCol + str(i) + ")/" + totalCostCol + str(i) + "),(365 / " + daysOwnedCol + str(i) + "))-1,0)"
        # print(formula)
        worksheet.write_formula(currentA,formula)


    # print(currentA)
    # workbook.close()

    entryFields = [ "entryDate", "entryType", "entrySecurity", "entrySymbol", "entryDescription","entryAmount","entryRemainingShares","entryPricePerShare" ]

    number_fmt = workbook.add_format()
    number_fmt.set_num_format('0.0000')
    date_fmt = workbook.add_format()
    date_fmt.set_num_format('mm/dd/yyyy')

    for key, value in symbols.items():
        # print('Symbol: ' + key)

        t = symbols[key]

        if t.numShares() <= 0:
            continue

        worksheetName = t.worksheetName()
        if len(worksheetName) > 30 :
            worksheetName = worksheetName[:30]

        worksheet = workbook.add_worksheet(worksheetName)

        worksheet.write(0,0,"Account")
        myColumnNum = 1
        for ef in entryFields:
            # myColumn = xlsxwriter.utility.xl_col_to_name(myColumnNum)
            currentA = myColumn + "1"
            worksheet.write(0,myColumnNum,ef)
            myColumnNum = myColumnNum +1


        # print(str(t))
        myColumnNum = 0
        myRow = 1
        for key, acct in t.accounts.items():

            # print(str(acct))

            for e in acct.entries:

                # print("e:",str(e))

                # myColumn = "A" + str(myRow)
                worksheet.write(myRow,0, acct.Name())

                myColumnNum = 1
                for ef in entryFields:
                    myColumn = xlsxwriter.utility.xl_col_to_name(myColumnNum)
                    currentA = myColumn + str(myRow)
                    if ef == 'entryAmount' or ef == 'entryRemainingShares' or ef == 'entryPricePerShare':
                        #
                        try:
                            num = float(e.Field(ef))
                        except:
                            num = 0.00
                        worksheet.write_number(myRow, myColumnNum, num ,number_fmt )
                    elif ef == 'entryDate':
                        format_str = '%m/%d/%Y'  # The format
                        datetime_obj = datetime.datetime.strptime(e.Field(ef), format_str)
                        worksheet.write_datetime(myRow,myColumnNum,datetime_obj,date_fmt)
                    else:
                        worksheet.write(myRow,myColumnNum,e.Field(ef))

                    myColumnNum = myColumnNum + 1


                myRow = myRow + 1

    # print( 'key:' +  key  + 'value:' + str(acct))

    # print(currentA)
    workbook.close()
    # bdata.printSymbols(symbols)