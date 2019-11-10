#
# Copyright (c) 2018 Pearce Software Solutions. All rights reserved.
#
import sys
import csv
import json
import xlsxwriter
import transaction
import re
import math


#  This class stores all the formats.
#  TODO: Make more flexible, pass in colours, font sizes, text alignments.
#
class XLSFormats:

    # all the initialization to be done is here...
    def __init__(self, workbook=None, num=2):
        self.workbook = workbook

        # self.colors = ['#FFFFFF','#EEEEEE','#DDDDDD','#CCCCCC','#BBBBBB','#AAAAAA']
        self.colors = ['#D9E2F3', '#FFFFFF']  #

        if num > len(self.colors):
            self.numberShades = len(self.colors)
        else:
            self.numberShades = num

        # types of formats supported
        self.currency = []
        self.dates = []
        self.formulas = []
        self.numbers = []
        self.percents = []
        self.timestamps = []
        self.texts = []
        self.urls = []

        if workbook != None:
            self.header_format = workbook.add_format()
            self.header_format.set_font_size(14)
            self.header_format.set_bold()
            #
            self.header_format.set_bg_color('#4674C1')
            self.header_format.set_align('center')
            self.header_format.set_align('vcenter')
            self.header_format.set_text_wrap()

            for i in range(self.numberShades):
                currency_format = workbook.add_format()
                currency_format.set_font_size(14)
                currency_format.set_num_format(8)
                currency_format.set_bg_color(self.colors[i])
                self.currency.append(currency_format)

                date_format = workbook.add_format()
                date_format.set_font_size(14)
                date_format.set_num_format('mm/dd/yyyy')
                date_format.set_bg_color(self.colors[i])
                self.dates.append(date_format)

                formula_format = workbook.add_format()
                formula_format.set_font_size(14)
                formula_format.set_bg_color(self.colors[i])
                self.formulas.append(formula_format)

                number_format = workbook.add_format()
                number_format.set_font_size(14)
                number_format.set_num_format(4)
                number_format.set_bg_color(self.colors[i])
                self.numbers.append(number_format)

                percent_fmt = workbook.add_format({'num_format': '0.00%'})
                percent_fmt.set_font_size(14)
                percent_fmt.set_bg_color(self.colors[i])
                self.percents.append(percent_fmt)

                timestamp_format = workbook.add_format()
                timestamp_format.set_font_size(14)
                timestamp_format.set_num_format('mm/dd/yyyy hh:mm:ss')
                timestamp_format.set_bg_color(self.colors[i])
                self.timestamps.append(timestamp_format)

                text_format = workbook.add_format()
                text_format.set_font_size(14)
                text_format.set_bg_color(self.colors[i])
                self.texts.append(text_format)

        else:
            self.header_format = None

    def headerFormat(self):
        return self.header_format

    def currencyFormat(self, row=1):
        i = row % self.numberShades
        return self.currency[i]

    def dateFormat(self, row=1):
        i = row % self.numberShades
        return self.dates[i]

    def formulaFormat(self, row=1):
        i = row % self.numberShades
        return self.formulas[i]

    def numberFormat(self, row=1):
        i = row % self.numberShades
        return self.numbers[i]

    def percentFormat(self, row=1):
        i = row % self.numberShades
        return self.percents[i]

    def timestampFormat(self, row=1):
        i = row % self.numberShades
        return self.timestamps[i]

    def textFormat(self, row=1):
        i = row % self.numberShades
        return self.texts[i]


#
#  Like the type to a format.
class TypeToFormat:
    def __init__(self, type, format):
        self.type = type
        self.format = format


# this class stores the tag return from the IEX format, The Label it will appear in the worksheet, it's type and format.
class IEXField:
    def __init__(self, tag, label, type, format=None):
        self.tag = tag
        self.label = label
        self.type = type
        self.format = format

    # used for debugging:
    def __str__(self):
        return self.tag + "," + self.label + "," + self.type


#  FFU:  This is for things like quote_data, stats_data, news_data
class IEXDataSet:
    def __init__(self, name):
        self.name = name

#
# TODO:  A class for a set of ColumnInfo class
#
#  This stores the information about a column
class ColumnInfo:
    def __init__(self, Worksheet, Name=None, Colnum=0, Height=1, InitSize=1, Maxsize=40):  # , doSplit = False):
        self.worksheet = Worksheet
        self.name = Name
        self.height = Height
        self.columnNumber = Colnum
        self.max_size = Maxsize
        self.size = InitSize
        # self.doSplit = doSplit # tell the ColumnInfo to split strings and find their individual word lenth.

    def __str__(self):
        return json.dumps(self.json())

    def json(self):
        blob = {}
        blob["Name"] = self.name
        blob["Height"] = self.height
        blob["ColumnNumber"] = self.columnNumber
        blob["MaxSize"] = self.max_size
        blob["Size"] = self.size

        if hasattr(self, "worksheet"):
            blob["Worksheet"] = str(self.worksheet)
        else:
            blob["Worksheet"] = str(None)

        return blob;

    def columnSize(self, factor=1.1):
        return self.size * factor

    def columnSetSize(self, factor=1.1):
        self.worksheet.set_column(self.columnNumber, self.columnNumber, self.columnSize(factor))

    def convertFloat(self, Value):
        if isinstance(Value, float) == False:
            try:
                fValue = float(Value)
                return fValue
            except:
                # write nothing
                print(self.name, " something is there, but its not a float[", Value, "]")
                # Worksheet.write(Row, Column, Value, Format)
                return None
        else:
            return Value

    #
    # Replacing these two lines with one call
    #                 worksheet.write(myRow, myCol, str(value), formats.textFormat(myRow))
    #                 ci.columnComputeSize(str(value), 'text')
    # TODO:  The size of the formula column is based on the size of the formula, not the result.  Need to base on size of result.
    def columnWrite(self, Row, Column, Value, Type=None, Format=None, Split=False, TheURL = None):

        if hasattr(self, "worksheet") == False:
            print("Error: No worksheet found")
            return

        if self.worksheet == None:
            print("Error: No worksheet assigned")
            return

        if Value == None or Value == "":
            # print(self.name," empty or blank Value")
            self.worksheet.write(Row, Column, "", Format)
            return

        if Type == 'currency' or Type == 'percent' or Type == 'number':
            fValue = self.convertFloat(Value)
            if fValue == None:
                # print("something is there, but its not a float")
                self.worksheet.write(Row, Column, Value, Format)
                self.columnComputeSize(Value)
                return

            self.worksheet.write_number(Row, Column, fValue, Format)
            self.columnComputeSize(fValue, Type)

        # elif Type == 'formula':
        #    print(Type)
        # elif Type == 'date':
        #    print(Type)
        # elif Type == 'timestamp':
        #    print(Type)
        elif Type == 'url':
            #
            self.worksheet.write_url(Row,Column,TheURL,Format,Value,None)
        else:
            # print("default type[",Type,"]")
            self.worksheet.write(Row, Column, Value, Format)
            self.columnComputeSize(Value, Type, Split)

    def columnComputeSize(self, field, type=None, doSplit=False):

        # Already hit max
        if self.size >= self.max_size:
            return

        mySize = 1
        if type == 'currency' and isinstance(field, float):
            # print(self.name,":",type)
            mySize = self.columnCurrencySize(field, 2, True)
        elif type == 'percent' and isinstance(field, float):
            mySize = self.columnPercentSize(field, 2, False)
        elif type == 'number' and isinstance(field, float):
            mySize = self.columnFloatSize(field)
        elif type == 'int' and isinstance(field, int):
            mySize = self.columnStringSize(field, doSplit)
        else:
            mySize = self.columnStringSize(field, doSplit)

        if mySize >= self.max_size:
            self.size = self.max_size

        elif mySize > self.size:
            self.size = mySize


    def columnFloatSize(self, f=0.00, decimalPlaces=2):

        # Already hit max
        if self.size >= self.max_size:
            return

        if isinstance(f, float) == False:
            sValue = str(f)
            # print("Error received non float value[",sValue,"] length[",len(sValue),"]")
            return len(sValue)

        myLen = 1 + decimalPlaces  # for the decimal point and decimal places
        # print("initial len:",myLen," for ", f)

        n = int(f)
        if n > 0:
            digits = int(math.log10(n)) + 1
        elif n == 0:
            digits = 1
        else:
            digits = int(math.log10(-n)) + 2  # +1 if you don't count the '-'

        num_commas = int((digits - 1) / 3)

        myLen = myLen + digits + num_commas

        return myLen

    def columnCurrencySize(self, f=0.0, decimalPlaces=2, parens=False):

        myLen = self.columnFloatSize(f, decimalPlaces)
        if f < 0.00 and parens == True:
            return myLen + 3  # allowing for ($xx.xx)
        else:
            return myLen + 1  # just add a $

    def columnPercentSize(self, f=0.0, decimalPlaces=2, parens=False):

        # expecting a perctage such as .0525.
        # excel will convert it to be 5.25%,
        # multiply by 100 to get the length.
        # Note when reading in a percentage you will need to divide by 100 before this function.
        myLen = self.columnFloatSize(f * 100, decimalPlaces)

        # a value of .05  will be represented as 5%
        # a value of 1.00 will be represented as 100%.
        if f < 0.00 and parens == True:
            return myLen + 3  # allowing for (x.xx%)
        elif f < 0.00:
            return myLen + 2  # allow for -xx%
        else:
            return myLen + 1  # just add a %

    def columnStringSize(self, p=None, doSplit=False):
        # check to see if the max has already been set

        if p == None or p == "":
            return 0

        mySize = 0
        # print("value:",p)
        # make sure we're dealing with the size of the string aka it's length.
        if isinstance(p, str):
            # print(p, "is a string")
            if doSplit:

                words = re.split('[ /]', p)
                # print(self.name," words:",words)
                if len(words) > self.height:
                    self.height = len(words)

                for w in words:
                    if len(w) > mySize:
                        mySize = len(w)
            else:
                mySize = len(p)
        else:
            mySize = len(str(p))

        return mySize


def InitType(formats):
    typeFormats = {}

    for name in ['currency', 'date', 'formula', 'number', 'percent', 'timestamp', 'text', 'url']:
        # print (name)
        if name == 'currency' and hasattr(formats, 'currencyFormat'):
            typeFormats[name] = formats.currencyFormat
        elif name == 'date' and hasattr(formats, 'dateFormat'):
            typeFormats[name] = formats.dateFormat
        elif name == 'formula' and hasattr(formats, 'formulaFormat'):
            typeFormats[name] = formats.formulaFormat
        elif name == 'number' and hasattr(formats, 'numberFormat'):
            typeFormats[name] = formats.numberFormat
        elif name == 'percent' and hasattr(formats, 'percentFormat'):
            typeFormats[name] = formats.percentFormat
        elif name == 'text' and hasattr(formats, 'textFormat'):
            typeFormats[name] = formats.textFormat
        elif name == 'timestamp' and hasattr(formats, 'timestampFormat'):
            typeFormats[name] = formats.timestampFormat
        elif name == 'url' and hasattr(formats, 'textFormat'):
            typeFormats[name] = formats.textFormat
        else:
            print("What the what? -->", name)

    return typeFormats


#
# this loads the data to fill in the IEXFields
def loadDataLabels(filename, formats):
    dataSet = []
    lookupReader = csv.reader(open(filename, newline=''), delimiter=',', quotechar='"')

    i = 0
    for row in lookupReader:
        i = i + 1
        if len(row) != 3:
            print("Error with row(", i, ") :", row)
            continue
        myType = row[2]
        myFormat = formats.get(myType)
        myField = IEXField(row[0], row[1], myType, myFormat)
        dataSet.append(myField)

    # for d in dataSet:
    #   print(d)

    return dataSet


#
# local run
if __name__ == "__main__":

    lookUps = dict()
    transactions = []
    inFilename = "transactions.csv"
    outFilename = "column_transactions.xlsx"

    i = 0
    for i in range(1, len(sys.argv)):
        if i == 1:
            inFilename = sys.argv[i]
        elif i == 2:
            outFilename = sys.argv[i]
        else:
            print("Ignoring extra arguments", sys.argv[i])

    workbook = xlsxwriter.Workbook(outFilename)
    formats = XLSFormats(workbook)

    #
    # test_ci = ColumnInfo("test", 0, 1, 1, 100)
    # for word in ["Test", "Testing Info", "Lets test this", "This_is_a_very_long_string", 1000000]:
    #     test_ci.columnStringSize(word, True)
    #     print(str(test_ci))
    #
    # worksheet = workbook.add_worksheet("test")
    # worksheet.write(0, test_ci.columnNumber, "This is a very long string")
    # worksheet.set_column(test_ci.columnNumber, test_ci.columnNumber, test_ci.size)

    T = transaction.Transactions(workbook, formats)
    T.loadTransactions(inFilename, lookUps)

    tags = transaction.TransactionJsonTags()
    hdrs = transaction.TransactionHeaders()

    colInfo = []

    worksheet = workbook.add_worksheet("Transactions")
    myCol = 0
    for h in hdrs:
        ci = ColumnInfo(worksheet, h, myCol)
        colInfo.append(ci)
        ci.columnWrite(0, myCol, h, 'text', formats.headerFormat())
        myCol = myCol + 1

    myRow = 1
    for t in T.transactions:
        myCol = 0
        # print(t)
        for h in hdrs:

            tag = tags[myCol]
            value = t.get_value(tag)

            ci = colInfo[myCol]
            if tag == 'year':
                formula = "=YEAR(A" + str(myRow + 1) + ")"
                ci.columnWrite(myRow, myCol, formula, 'formula', formats.formulaFormat(myRow))

            elif tag == 'month':
                formula = "=MONTH(a" + str(myRow + 1) + ")"
                ci.columnWrite(myRow, myCol, formula, 'formula', formats.formulaFormat(myRow))

            elif tag == "shares":
                ci.columnWrite(myRow, myCol, value, 'number', formats.numberFormat(myRow))

            elif tag == "invest_amt" or tag == "amount":
                # print(tag,":",value)
                ci.columnWrite(myRow, myCol, value, 'currency', formats.currencyFormat(myRow))

            else:
                ci.columnWrite(myRow, myCol, value, 'text', formats.textFormat(myRow))

            myCol = myCol + 1

        myRow = myRow + 1

    #
    # go back and compute set the column sizes
    # myCol = 0
    for ci in colInfo:
        print(str(ci))
        ci.columnSetSize(1.2)

    workbook.close()
