
import csv
import xlsxwriter

def PortfolioHeaders():
    return [ "Name","Symbol","Quote","Price Day Change","Price Day Change (%)","Shares","Cost Basis","Market Value","Gain/Loss","Gain/Loss (%)"]

def PortfolioLabels():
    return ["name","symbol","quote","price_day_change","price_day_change_pct","shares","cost_basis","market_value","gain_loss","gain_loss_pct"]


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
        if len(row) == 10:
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
                    print("Adding symbol:", row[1])

                row[0] = sname
                datum = dict()
                for i in range(10):
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
def WritePortfolioValueWorksheet(pvalue, workbook):
    if isinstance(pvalue,dict) == False:
        print("Invalid argument pvalue")
        return

    worksheet = workbook.add_worksheet('Portfolio')
    money_fmt = workbook.add_format({'num_format': '$#,##0.00'})
    percent_fmt = workbook.add_format({'num_format': '0.00%'})

    myKeys = pvalue.keys()
    # myKeys.sort()
    labels = PortfolioLabels()
    headers = PortfolioHeaders()

    myRow = 0
    myColumn = 0
    for h in headers:
        worksheet.write(myRow,myColumn,h)
        myColumn = myColumn + 1

    myRow = 1
    for k in myKeys:
        v = pvalue[k]
        myColumn = 0
        for l in labels:
            if l == "quote" or l == "price_day_change" or l == "shares":

                worksheet.write_number(myRow, myColumn, float(v[l]))
            elif l == "cost_basis" or l == "market_value" or l == "gain_loss":
                worksheet.write_number(myRow, myColumn, float(v[l]), money_fmt )
            elif l == "price_day_change_pct" or l == "gain_loss_pct":
                worksheet.write_number(myRow, myColumn, float(v[l])/100, percent_fmt)
            else:
                worksheet.write(myRow,myColumn,v[l])

            myColumn = myColumn + 1
        myRow = myRow + 1

if __name__ == "__main__":
    pvalue = dict()
    name = "KP2019 - Investing - Portfolio Value - Group by Security - 2019-06-29.csv"
    workbook = xlsxwriter.Workbook("portfolios.xlsx")

    LoadPortfolioValue(name,pvalue, None)
    WritePortfolioValueWorksheet(pvalue,workbook)
    workbook.close()

    print(pvalue)
