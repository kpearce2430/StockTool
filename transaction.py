
import sys
import csv
import xlsxwriter
import json
import common_xls_formats

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

class Transactions:

    def __init__(self, workbook, formats ):

        self.workbook = workbook
        self.formats = formats
        self.transactions = []
        self.jsonTags = []
        self.headers = []


    def loadTransactions(self,infilename, lookups ):

        if isinstance(lookups, dict) == False:
            print("Invalid argument lookups")
            return

        transReader = csv.reader(open(infilename, newline=''), delimiter=',', quotechar='"')

        i = 0
        securityRow = -1
        symbolRow = -1
        myTags = []
        # myHeaders = []

        for row in transReader:

            i = i + 1

            if len(row) < 3:
                continue

            # The first two columns from Quicken 18 are junk.
            # Get rid of them.
            del row[0]
            del row[0]

            # Skip the hard coded Date line.
            if row[0] == 'Date' and len(self.headers) == 0 and len(self.jsonTags) == 0:
                j = 0
                for r in row:
                    self.headers.append(r)
                    tag = returnTag(r)
                    self.jsonTags.append(tag)
                    if r == "Security":
                        securityRow = j
                    if r == "Symbol":
                        symbolRow = j

                    j = j + 1

                print("# Tags:", len(self.jsonTags))
                print("# Headers:", len(self.headers))
                print("Security Row:", securityRow)
                print("Symbol Row:", symbolRow)
                continue


            if len(self.jsonTags) <= 0:
                # havent found the header row yet. lets just...
                continue


            if (len(row) != len(self.headers)):
                continue

            info = Transaction()

            for j in range(len(self.jsonTags)):

                if self.jsonTags[j] == 'None':
                    continue

                # Get the symbols from the lookup
                if self.headers[j] == 'Security':
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

                    if row[symbolRow] == None or row[symbolRow] == "":
                        row[symbolRow] = value


                if self.headers[j] == 'Shares' or self.headers[j] == 'Amount' or self.headers[j] == 'Invest Amt':
                    # Remove the commas
                    row[j] = row[j].replace(',', '')

                info.set_value(self.jsonTags[j],row[j])
                j = j + 1

            i = i + 1

            self.transactions.append(info)

        self.headers.append("Year")
        self.jsonTags.append("year")
        self.headers.append("Month")
        self.jsonTags.append("month")

        print("Loaded ", len(self.transactions), " transactions")

    def writeTransactionWorksheetHeaders(self, worksheet, startRow = 0, startColumn = 0 ):
        myCol = startColumn
        # headers = TransactionHeaders()
        for h in self.headers:
            worksheet.write(startRow, myCol, h, self.formats.headerFormat())
            myCol = myCol + 1

    def writeTransactionWorksheetRow(self, worksheet, trans, startRow = 0, startColumn = 0):

        myCol = startColumn

        for i in range(len(self.headers)):

            myTag = self.jsonTags[i]
            value = trans.get_value(myTag)

            if myTag == 'date':
                worksheet.write(startRow, myCol, value, self.formats.dateFormat(startRow))

            elif myTag == 'shares':
                try:
                    worksheet.write_number(startRow, myCol, float(value),self.formats.numberFormat(startRow))
                except:
                    worksheet.write_number(startRow, myCol, 0.00,self.formats.numberFormat(startRow))

            elif myTag == 'amount':

                if trans.get_value("type") == "Reinvest Dividend":
                    value = trans.get_value("invest_amt")

                try:
                    worksheet.write_number(startRow, myCol, float(value),self.formats.currencyFormat(startRow))
                except:
                    print(myTag, "Not a number(", str(value), ")")

            elif myTag == 'invest_amt':
                value = trans.get_value("invest_amt")
                try:
                    worksheet.write_number(startRow, myCol, float(value),self.formats.currencyFormat(startRow))
                except:
                    worksheet.write_number(startRow, myCol, 0.00,self.formats.currencyFormat(startRow))
                    # print(myTag, "Not a number(", str(value), ")")

            elif myTag == 'year':
                formula = "=YEAR(" + "A" + str(startRow + 1) + ")"
                worksheet.write_formula(startRow, myCol, formula, self.formats.formulaFormat(startRow))

            elif myTag == 'month':
                formula = "=MONTH(" + "A" + str(startRow + 1) + ")"
                worksheet.write_formula(startRow, myCol, formula, self.formats.formulaFormat(startRow))

            else:
                worksheet.write(startRow, myCol, value, self.formats.textFormat(startRow))

            myCol = myCol + 1

    def writeTransactions(self, startRow = 0, startColumn = 0, worksheet = None, filterFunc = None, fArgs = None ):

        if worksheet == None:
            if self.workbook == None:
                return None
            myWorksheet = self.workbook.add_worksheet('Transactions')
        else:
            myWorksheet = worksheet

        row = startRow
        column = startColumn
        self.writeTransactionWorksheetHeaders(myWorksheet,row,column)

        row = row + 1
        for t in self.transactions:
            if filterFunc != None:
                if filterFunc(t,fArgs) == True:
                    # print("writing")
                    self.writeTransactionWorksheetRow(myWorksheet, t, row, startColumn)
                    row = row + 1
            else:
                self.writeTransactionWorksheetRow(myWorksheet, t, row, startColumn)
                row = row + 1

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

def pickSymbol(trans,fArgs):
    mySymbol = trans.get_value("symbol")
    # print(fArgs,mySymbol)
    if mySymbol == fArgs:
        # print(fArgs, mySymbol,"True")
        return True
    else:
        # print(fArgs, mySymbol,"False")
        return False

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

    workbook = xlsxwriter.Workbook(outFilename)
    formats = common_xls_formats.XLSFormats(workbook)
    T = Transactions(workbook,formats);
    T.loadTransactions(inFilename,lookUps)
    T.writeTransactions(0,0)

    myWorksheet = workbook.add_worksheet('APPLE')
    T.writeTransactions(0,0,myWorksheet,pickSymbol,"AAPL")

    workbook.close()

