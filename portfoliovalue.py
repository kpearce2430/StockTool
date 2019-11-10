#
# Copyright (c) 2018 Pearce Software Solutions. All rights reserved.
#
import sys
import csv
import xlsxwriter
import common_xls_formats

def PortfolioHeaders():
    return [ "Name","Symbol","Quote","Price Day Change","Price Day Change (%)","Shares","Cost Basis","Market Value","Average Cost Per Share","Gain/Loss 12-Month","Gain/Loss","Gain/Loss (%)" ]
    # return [ "Name","Symbol","Quote","Price Day Change","Price Day Change (%)","Shares","Cost Basis","Market Value","Gain/Loss","Gain/Loss (%)"]

def PortfolioLabels():
    return ["name","symbol","quote","price_day_change","price_day_change_pct","shares","cost_basis","market_value","avg_cost_per_share","gain_loss_last_12m","gain_loss","gain_loss_pct"]


def LoadPortfolioValue(name, pvalue, lookups = None):

    if isinstance(pvalue,dict) == False:
        print("Invalid argument pvalue")
        return


    if lookups != None and isinstance(lookups,dict) == False:
        print("Invalid argument lookups")
        return

    labels = PortfolioLabels() #



    lookupReader = csv.reader(open(name,newline=''),delimiter=',', quotechar='"')
    for row in lookupReader:
        #
        if len(row) != 12:
            # print("Invalid row:", row)
            continue
        else:
            sname = row[0]
            if len(sname) == 1:
                continue

            # trim off the garbage character
            sname = sname[1:]

            if sname == "Cash":
                print(sname,":",row[7])
            elif sname == "Totals":
                print(sname,":",row[7])
            else:
                # look for a missing symbol
                if lookups != None and row[1] == '':  # symbol is blank
                    row[1] = lookups.get(sname)
                    # print("Adding symbol:", row[1])

                row[0] = sname
                datum = dict()
                for i in range(12):
                    if i > 1:
                        if row[i] == "Add":
                            value = "0"
                        else:
                            value = str(row[i])
                            value = value.replace(',', '')
                            value = value.replace('$', '')
                            value = value.replace('#', '')
                            value = value.replace('%', '')
                        # print(labels[i],":",value)
                        # datam[labels[i]] = value
                    else:
                        value = row[i]

                    datum[labels[i]] = value


                pvalue[row[1]] = datum
                # print(sname,":",datum)


#
#
#
def WritePortfolioValueWorksheet(pvalue, workbook, formats):

    if isinstance(pvalue,dict) == False:
        print("Invalid argument pvalue")
        return

    worksheet = workbook.add_worksheet('Portfolio')

    myKeys = pvalue.keys()
    # myKeys.sort()
    labels = PortfolioLabels()
    headers = PortfolioHeaders()
    columnInfo = []

    myRow = 0
    myColumn = 0
    for h in headers:
        ci = common_xls_formats.ColumnInfo(worksheet,h,myColumn)
        ci.columnWrite(myRow,myColumn,h,'text',formats.headerFormat(),True)
        columnInfo.append(ci)
        myColumn = myColumn + 1

    # for myCol in range(0,len(columnInfo)):
    #    ci = columnInfo[myCol]
    #    print(myCol,":",str(ci))

    myRow = 1
    for k in myKeys:
        v = pvalue[k]
        myColumn = 0
        for l in labels:
            ci = columnInfo[myColumn]

            if l == "shares":
                ci.columnWrite(myRow,myColumn,v[l],'number',formats.numberFormat(myRow))

            elif l == "quote" or l == "price_day_change" or l == 'market_value' or l == 'gain_loss' or l == 'avg_cost_per_share' or l == 'gain_loss_last_12m' or l == 'cost_basis':
                ci.columnWrite(myRow,myColumn,v[l],'currency',formats.currencyFormat(myRow))

            elif l == "price_day_change_pct" or l == "gain_loss_pct":
                # since I'm already reading in a percent, it needs to be converted back to a float
                fValue = ci.convertFloat(v[l])
                if fValue != None:
                    fValue = fValue / 100;
                    ci.columnWrite(myRow,myColumn,fValue,'percent',formats.percentFormat(myRow))
                else:
                    ci.columnWrite(myRow,myColumn,v[l],"text",formats.textFormat(myRow))

            else:
                ci.columnWrite(myRow,myColumn,v[l],'text',formats.textFormat(myRow))

            myColumn = myColumn + 1

        myRow = myRow + 1

    for ci in columnInfo:
        ci.columnSetSize(1.3)
        # print(myCol,":",str(ci))
        # worksheet.set_column(ci.columnNumber,ci.columnNumber,ci.columnSize(1.3))

if __name__ == "__main__":
    pvalue = dict()

    inFilename = "portfolio_value.csv"
    outFilename = "portfolio_value.xlsx"

    i = 0
    for i in range(1, len(sys.argv)):

        if i == 1:
            inFilename = sys.argv[i]
        elif i == 2:
            outFilename = sys.argv[i]
        else:
            print("Ignoring extra arguments",sys.argv[i])

    workbook = xlsxwriter.Workbook(outFilename)
    formats = common_xls_formats.XLSFormats(workbook,5)

    LoadPortfolioValue(inFilename,pvalue, None)
    WritePortfolioValueWorksheet(pvalue,workbook,formats)
    workbook.close()

    # print(pvalue)
