
import csv
import xlsxwriter

def TransactionHeaders():

    return ["Date","Type","Security","Security/Payee","Description","Shares","Amount","Account","Year","Month"]


def LoadTransactions(infilename, transactions, lookups ):

    if isinstance(transactions,list) == False:
        print("Invalid argument transactions")
        return

    if isinstance(lookups, dict) == False:
        print("Invalid argument lookups")
        return

    transReader = csv.reader(open(infilename, newline=''), delimiter=',', quotechar='"')

    i = 0

    for row in transReader:

        i = i + 1
        if (len(row) < 10):
            print("skipping[", i, "] items[", len(row), "] [", row, "]")
            continue

        # The first two columns from Quicken 18 are junk.
        # Get rid of them.
        del row[0]
        del row[0]

        # Skip the hard coded Date line.
        if row[0] == 'Date':
            print("skipping[",i,"item[",row[2],"] [",row,"]")
            continue

        # 3 - Get the symbols from the lookup
        value = lookups.get(row[3])

        if value == None:
            value = "Missing"

        row[3] = value

        # 6 Shares - remove the commas
        row[5] = row[5].replace(',', '')

        # 7 Amount - remove the commas
        row[6] = row[6].replace(',', '')

        i = i + 1

        transactions.append(row)

    print("Loaded ", len(transactions), " transactions")

def WriteTransactionWorksheet( transactions, workbook):

    if isinstance(transactions, list) == False:
       print("Invalid argument pvalue")
       return

    date_fmt = workbook.add_format()
    date_fmt.set_num_format('mm/dd/yyyy')

    worksheet = workbook.add_worksheet('Transactions')

    row = 0
    column = 0
    headers = TransactionHeaders()
    for h in headers:
        worksheet.write(row,column,h)
        column = column + 1


    row = row + 1
    for data in transactions:

        if isinstance(data,list) == False:
            print("Invalid row in transaction list")
            return

        column = 0;
        for item in data:

            if column == 0:
                worksheet.write(row,column,item,date_fmt)
            elif column == 5 or column == 6:
                if item == "":
                    item = "0.00"

                worksheet.write_number(row,column,float(item))
            else:
                worksheet.write(row,column,item)

            column = column + 1

        formula = "=YEAR(" + "A"+str(row+1) + ")"
        worksheet.write_formula(row,column,formula)

        column = column + 1

        formula = "=MONTH(" + "A" + str(row+1) + ")"
        worksheet.write_formula(row, column, formula)

        row = row + 1

if __name__ == "__main__":
    lookUps = dict()
    transactions = []
    filename = "KP2019-export-2019-06-29.csv"
    LoadTransactions(filename, transactions, lookUps)

    workbook = xlsxwriter.Workbook("transactions.xlsx")
    WriteTransactionWorksheet(transactions, workbook)
    workbook.close()



