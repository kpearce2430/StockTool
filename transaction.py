
import csv
import xlsxwriter
import json

class Transaction:

    def __init__(self, entry=None):
        # self.name = ""
        self.info = dict()

        if isinstance(entry,dict):
            for k in entry:
                self.info[k] = entry[k]
        

    def __str__(self):
        return "Transaction:" + json.dumps(self.info)


    def set_value(self,key=None,value=None):
        
        if key != None:
            self.info[key] = value
        else:
            print("Error - Missing key in Transaction.set_value")
        
    def get_value(self, key):
        if key != None:
            return self.info.get(key)
        else:
            return None
    
    def  getDate(self):
        return self.get_value('date')

def TransactionJsonTags():
    return ["date", "type", "security","security_payee","description","shares","amount","account","year","month"]

def TransactionHeaders():
    return ["Date","Type","Security","Security/Payee","Description/Category","Shares","Amount","Account","Year","Month"]


def returnTag(hdr):
    headers = TransactionHeaders()
    tags = TransactionJsonTags()

    try:
        idx = headers.index(hdr)
    except:
        print("No element found for:",hdr)
        return "None"

    if idx >= len(tags):
        print("Error:  Tag length mismatch")

    return tags[idx]

def LoadTransactions(infilename, transactions, lookups ):

    if isinstance(transactions,list) == False:
        print("Invalid argument transactions")
        return

    if isinstance(lookups, dict) == False:
        print("Invalid argument lookups")
        return

    transReader = csv.reader(open(infilename, newline=''), delimiter=',', quotechar='"')

    i = 0

    myTags = []
    myHeaders = []
    securityRow = -1
    for row in transReader:

        i = i + 1

        if len(row) < 3:
            # print("skipping[", i, "] [", row, "]")
            continue

        # The first two columns from Quicken 18 are junk.
        # Get rid of them.
        del row[0]
        del row[0]

        # Skip the hard coded Date line.
        if row[0] == 'Date' and len(myHeaders) == 0 and len(myTags) == 0:
            j = 0
            for r in row:
                myHeaders.append(r)
                tag = returnTag(r)
                # print(r,":",tag)
                myTags.append(tag)
                if r == "Security":
                    securityRow = j
                j = j + 1

            print("# Tags:", len(myTags))
            print("# Headers:", len(myHeaders))
            print("Security Row:", securityRow)
            # print("skipping[",i,"item[",row[2],"] [",row,"]")
            continue


        if len(myTags) <= 0:
            # havent found the header row yet. lets just...
            continue


        if (len(row) != len(myTags)):
            # print("skipping[", i, "] items[", len(row), "] [", row, "]")
            continue

        info = Transaction()

        for j in range(len(myTags)):
            # print(j)

            if myTags[j] == 'None':
                continue

            # Get the symbols from the lookup
            if myHeaders[j] == 'Security':
                value = lookups.get(row[j])
                if value == None:
                    value = "Missing"
                row[j] = value

            if myHeaders[j] == 'Shares' or myHeaders[j] == 'Amount':
                # Femove the commas
                row[j] = row[j].replace(',', '')

            info.set_value(myTags[j],row[j])
            j = j + 1

        # print(info)
        i = i + 1

        transactions.append(info)

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

    tags = TransactionJsonTags()

    for data in transactions:

        column = 0
        for i in range(len(headers)):
            myTag = tags[i]

            value = data.get_value(myTag)

            if myTag == 'date':
                worksheet.write(row, column, value, date_fmt)


            elif myTag == 'shares' or myTag == 'amount':
                if value.isnumeric():
                    worksheet.write_number(row, column, float(value))
                else:
                    worksheet.write(row, column, value)

            elif myTag == 'year':
                formula = "=YEAR(" + "A" + str(row + 1) + ")"
                worksheet.write_formula(row, column, formula)

            elif myTag == 'month':
                formula = "=MONTH(" + "A" + str(row + 1) + ")"
                worksheet.write_formula(row, column, formula)

            else:
                worksheet.write(row,column,value)

            column = column + 1

        row = row + 1

if __name__ == "__main__":
    lookUps = dict()
    transactions = []
    filename = "KP2019-export-2019-09-01.csv"
    LoadTransactions(filename, transactions, lookUps)

    workbook = xlsxwriter.Workbook("transactions.xlsx")
    WriteTransactionWorksheet(transactions, workbook)
    workbook.close()

