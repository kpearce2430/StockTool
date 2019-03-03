
import csv
# import json
import xlsxwriter
import basedatacsv as bdata
import urllib3

def LoadLookup(name, lookup):
    lookupReader = csv.reader(open(name,newline=''),delimiter=',', quotechar='"')
    for row in lookupReader:
        if len(row) == 2:
            # print(row[0],row[1])
            lookup[row[0]] = row[1]
        else:
            print("huh:",row)


if __name__ == "__main__":

    # load up the lookup table
    lookUps = dict()
    LoadLookup('lookup.csv', lookUps)
    # print("lookups",len(lookUps))

    http = urllib3.PoolManager()
    urllib3.disable_warnings()

    symbols = dict()
    unique_accounts = []


    # prepare the worksheet
    basename = 'KP2013-export-2019-03-02'
    infilename = basename + '.csv'
    outfilename = basename + '.xlsx'

    workbook = xlsxwriter.Workbook(outfilename)
    worksheet = workbook.add_worksheet('Lookups')
    myRow = 1

    myKeys = lookUps.keys()
    for k in myKeys:
        v = lookUps[k]
        currentA = 'A' + str(myRow)
        currentB = 'B' + str(myRow)
        worksheet.write(currentA,k)
        worksheet.write(currentB,v)
        myRow = myRow+1

    worksheet = workbook.add_worksheet('Transactions')


    # Transasction headers
    worksheet.write('A1','Date')
    worksheet.write('B1','Type')
    worksheet.write('C1','Security')
    worksheet.write('D1','Security/Payee')
    worksheet.write('E1','Description')
    worksheet.write('F1','Shares')
    worksheet.write('G1','Amount')
    worksheet.write('H1','Account')
    worksheet.write('I1','Year')
    worksheet.write('J1','Month')

    i = 0
    xRow = 2
    transReader = csv.reader(open(infilename, newline=''), delimiter=',', quotechar='"')

    # A - Date
    # B - Type
    # C - Security
    # D - Security/Payee
    # E - Description
    # F - Shares
    # G - Amount
    # H - Account
    # I - Year
    # J - Month
    TheColumns = [ "A", "B", "C", "D", "E","F","G","H"]
    # TheRows = [ 2, 3, 4, 5, 6, 7 ]
    for row in transReader:

        if (len(row) < 10):
            print("skipping[", i, "] items[", len(row), "] [", row, "]")
            continue

        # The first two columns from Quicken 18 are junk.
        # Get rid of them.
        del row[0]
        del row[0]

        i = i + 1

        if row[0] == 'Date':
            print("skipping[",i,"item[",row[2],"] [",row,"]")
            continue

        j = 0

        value = lookUps.get(row[3])

        if value == None:
           value = "Missing"
           # print("Missing",row[5])

        # if value != 'CDC':
        #     continue;
        # else:
        #     print(row)

        row[3] = value

        for c in TheColumns:

            currentA = c  + str(xRow)
            if c == "F" or c == "G":
               try:
                    row[j] = row[j].replace(',','')
                    val = float(row[j])
                    worksheet.write(currentA,float(row[j]))
               except ValueError:
                   worksheet.write(currentA, row[j])
            else:
                worksheet.write(currentA,row[j])

            j = j + 1

        currentA = "I" + str(xRow)
        formula = "=YEAR(" + "A"+str(xRow) + ")"
        worksheet.write_formula(currentA,formula)

        currentA = "J" + str(xRow)
        formula = "=MONTH(" + "A"+str(xRow) + ")"
        worksheet.write_formula(currentA,formula)

        xRow = xRow + 1
        # add row to the master list
        bdata.ProcessRow(row,symbols,unique_accounts,http)


    unique_accounts.sort()

    # bdata.printSymbols( symbols )

    details = []
    bdata.createSheet(symbols, unique_accounts, details)

    worksheet = workbook.add_worksheet('Stock Analysis')

    # First 2 columns are Name and Symbol
    myKeys = []
    worksheet.write('A1','Name')
    myKeys.append('Name')
    worksheet.write('B1','Symbol')
    myKeys.append('Symbol')

    # Add in the individual accounts next
    myOrd = ord('C')
    for a in unique_accounts:
        currentA = str(chr(myOrd)) + "1"
        worksheet.write(currentA,a)
        myKeys.append(a)
        myOrd = myOrd + 1

    # add in Total Shares
    myKeys.append('Total Shares')
    currentA = str(chr(myOrd)) + "1"
    worksheet.write(currentA, 'Total Shares')
    totalSharesCol = str(chr(myOrd))

    myOrd = myOrd + 1

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
                currentA = str(chr(myOrd)) + "1"
                worksheet.write(currentA, k)

                if k == 'Latest Price':
                    latestPriceCol = str(chr(myOrd))
                elif k == 'Total Value':
                    totalValueCol = str(chr(myOrd))
                elif k == 'Total Cost':
                    totalCostCol = str(chr(myOrd))
                elif k == 'Yearly Dividend':
                    yearlyDividendCol = str(chr(myOrd))
                elif k == 'Days Owned':
                    daysOwnedCol = str(chr(myOrd))
                elif k == 'Dividends Received':
                    dividendsRecCol = str(chr(myOrd))

                myOrd = myOrd + 1
                # print("adding ",k)


    # push it out to the worksheet
    xRow = 2
    for d in details:
        myOrd = ord('A')

        if d['Latest Price'] == 0:
            # print("Skipping ",d["Name"])
            continue

        for l in myKeys:
            v = d[l]
            currentA = str(chr(myOrd)) + str(xRow)
            worksheet.write(currentA,v)
            myOrd = myOrd + 1

        xRow = xRow + 1

    stockOrd = myOrd
    # worksheet = workbook.add_worksheet('Fund Analysis')

    # xRow = 2
    # myOrd = ord('A')
    # for k in myKeys:
    #    currentA = str(chr(myOrd)) + "1"
    #    worksheet.write(currentA,k)
    #    myOrd = myOrd + 1

    for d in details:
        myOrd = ord('A')

        if d['Latest Price'] != 0:
            # print("Skipping",d['Name'])
            continue

        for l in myKeys:
            v = d[l]
            currentA = str(chr(myOrd)) + str(xRow)

            if l == 'Total Value':
                formula = "=" + totalSharesCol + str(xRow) + "*" + latestPriceCol + str(xRow)
                # print(formula)
                worksheet.write_formula(currentA,formula)
            elif l == 'Net':
                formula = "=" + totalValueCol + str(xRow) + "-" + totalCostCol + str(xRow)
                worksheet.write_formula(currentA, formula)
            else:
                worksheet.write(currentA, v)

            myOrd = myOrd + 1

        xRow = xRow + 1

    if myOrd == ord('A'):
        myOrd = stockOrd

    # add the formulas for
    # ... Percentage of the Portfolio

    percent_fmt = workbook.add_format({'num_format': '0.00%'})

    currentA = str(chr(myOrd)) + "1"
    worksheet.write(currentA,'Percentage Portfolio')
    currentA= str(chr(myOrd)) + ":" + str(chr(myOrd))
    worksheet.set_column( currentA, None, percent_fmt)
    # print(str(chr(myOrd)),xRow)

    currentA = str(chr(myOrd+1)) + "1"
    worksheet.write(currentA,'Dividend Yield')
    currentA= str(chr(myOrd+1)) + ":" + str(chr(myOrd+1))
    worksheet.set_column( currentA, None, percent_fmt)
    # print(str(chr(myOrd+1)),xRow)

    currentA = str(chr(myOrd+2)) + "1"
    worksheet.write(currentA,'ROI')
    currentA= str(chr(myOrd+2)) + ":" + str(chr(myOrd+2))
    worksheet.set_column( currentA, None, percent_fmt)
    # print(str(chr(myOrd+2)),xRow)

    currentA = str(chr(myOrd+3)) + "1"
    worksheet.write(currentA,'Annual Return')
    currentA= str(chr(myOrd+3)) + ":" + str(chr(myOrd+3))
    worksheet.set_column( currentA, None, percent_fmt)
    # print(str(chr(myOrd+3)),xRow)

    currentA = str(chr(myOrd+4)) + "1"
    worksheet.write(currentA,'CAGR')
    currentA= str(chr(myOrd+4)) + ":" + str(chr(myOrd+4))
    worksheet.set_column( currentA, None, percent_fmt)
    print(str(chr(myOrd+4)),xRow)

    for i in range(2,xRow):
        currentA =  str(chr(myOrd)) + str(i)
        currentB =  totalValueCol + str(i)
        formula = '=' + currentB + '/ (SUM($' + totalValueCol + "$2:$" + totalValueCol + "$" + str(xRow-1) + "))"
        # print(formula)
        worksheet.write_formula(currentA,formula)

        #IF(K48>0,R48/K48,0)
        currentA =  str(chr(myOrd+1)) + str(i)
        formula = "=if(" + yearlyDividendCol + str(i) + "> 0 ," + yearlyDividendCol + str(i) + "/" + latestPriceCol + str(i) + ",0)"
        worksheet.write_formula(currentA,formula)

        currentA = str(chr(myOrd+2)) + str(i)
        formula = "=if(" + totalCostCol + str(i) + ">0,(" + totalValueCol + str(i) + "-" + totalCostCol + str(i) + ") / " + totalCostCol + str(i) + ", 0)"
        worksheet.write_formula(currentA,formula)

        # =if(n2>0,POWER((L2/N2),(365/U2))-1,0)
        currentA = str(chr(myOrd+3)) + str(i)
        formula = "=if(" + totalCostCol + str(i) + ">0,power((" + totalValueCol + str(i) + "/" + totalCostCol + str(i) + "),(365 / " + daysOwnedCol + str(i) + "))-1, 0)"
        worksheet.write_formula(currentA,formula)

        # =IF(N37>0,POWER(((L37+M37)/N37),(365/U37))-1,0)
        currentA = str(chr(myOrd+4)) + str(i)
        formula = "=if(" + totalCostCol + str(i) + ">0,power(((" + totalValueCol + str(i) + "+" + dividendsRecCol + str(i) + ")/" + totalCostCol + str(i) + "),(365 / " + daysOwnedCol + str(i) + "))-1,0)"
        # print(formula)
        worksheet.write_formula(currentA,formula)


    print(chr(myOrd+4),myOrd+4)
    workbook.close()