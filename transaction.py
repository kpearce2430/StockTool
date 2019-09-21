
import sys
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
    return ["date", "type", "security","symbol","security_payee","description","shares","invest_amt","amount","account","year","month"]

def TransactionHeaders():
    return ["Date","Type","Security","Symbol","Security/Payee","Description/Category","Shares","Invest Amt","Amount","Account","Year","Month"]


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
    symbolRow = -1
    for row in transReader:

        i = i + 1

        if len(row) < 3:
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
                myTags.append(tag)
                if r == "Security":
                    securityRow = j
                if r == "Symbol":
                    symbolRow = j

                j = j + 1

            print("# Tags:", len(myTags))
            print("# Headers:", len(myHeaders))
            print("Security Row:", securityRow)
            continue


        if len(myTags) <= 0:
            # havent found the header row yet. lets just...
            continue


        if (len(row) != len(myTags)):
            continue

        info = Transaction()

        for j in range(len(myTags)):

            if myTags[j] == 'None':
                continue

            # Get the symbols from the lookup
            if myHeaders[j] == 'Security':
                # I found over the years that symbols occassionally get reused.  I use
                # lookups to override any symbols that Quicken may put in.  If there isn't
                # a symbol in lookup and symbols are provided in the sheet.  I use the
                # sheet.
                value = lookups.get(row[j])
                if value == None:
                    if symbolRow != -1:
                        value = row[symbolRow]
                    else:
                        value = "Missing"

                row[j] = value

            if myHeaders[j] == 'Shares' or myHeaders[j] == 'Amount' or myHeaders[j] == 'Invest Amt':
                # Femove the commas
                row[j] = row[j].replace(',', '')

            info.set_value(myTags[j],row[j])
            j = j + 1

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

            elif myTag == 'shares' or myTag == 'invest_amt':
                try:
                    worksheet.write_number(row, column, float(value))
                except:
                    worksheet.write_number(row, column, 0.00)

            elif myTag == 'amount':

                if data.get_value("type") == "Reinvest Dividend":
                    value = data.get_value("invest_amt")

                try:
                    worksheet.write_number(row, column, float(value))
                except:
                    print(myTag,"Not a number(", str(data),")")

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
    inFilename = "transactions.csv"
    outFilename = "transactions.xlsx"

    i = 0
    for i in range(1, len(sys.argv)):

        if i == 1:
            inFilename = sys.argv[i]
        elif i == 2:
            outFilename = sys.argv[i]
        else:
            print("Ignoring extra arguments",sys.argv[i])

    LoadTransactions(inFilename, transactions, lookUps)
    workbook = xlsxwriter.Workbook(outFilename)
    WriteTransactionWorksheet(transactions, workbook)
    workbook.close()

