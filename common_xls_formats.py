#
# Copyright (c) 2018 Pearce Software Solutions. All rights reserved.
#
#import sys
import csv
#import xlsxwriter
#import json

class XLSFormats:

    def __init__(self, workbook = None, num = 2 ):
        self.workbook = workbook

        # self.colors = ['#FFFFFF','#EEEEEE','#DDDDDD','#CCCCCC','#BBBBBB','#AAAAAA']
        self.colors = ['#D9E2F3','#FFFFFF'] #

        if num > len(self.colors):
            self.numberShades = len(self.colors)
        else:
            self.numberShades = num

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
            # self.header_format.set_bg_color('gray')
            self.header_format.set_bg_color('#4674C1')
            self.header_format.set_align('center')
            self.header_format.set_align('vcenter')

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

    def dateFormat(self,row = 1):
        i = row % self.numberShades
        return self.dates[i]

    def formulaFormat(self, row=1):
        i = row % self.numberShades
        return self.formulas[i]

    def numberFormat(self,row = 1):
        i = row % self.numberShades
        return self.numbers[i]

    def percentFormat(self, row=1):
        i = row % self.numberShades
        return self.percents[i]

    def timestampFormat(self,row = 1):
        i = row % self.numberShades
        return self.timestamps[i]

    def textFormat(self,row = 1):
        i = row % self.numberShades
        return self.texts[i]

#
#  currency
class TypeToFormat:
    def __init__(self, type, format):
        self.type = type
        self.format = format

class IEXField:
    def __init__(self, tag, label, type, format = None ):
        self.tag = tag
        self.label = label
        self.type = type
        self.format = format

    def __str__(self):
        return self.tag+","+self.label+","+self.type

#  This is for things like quote_data, stats_data, news_data
class IEXDataSet:
    def __init__(self, name ):
        self.name = name


def InitType(formats):
    typeFormats = {}

    for name in ['currency','date','formula','number','percent','timestamp','text','url']:
        # print (name)
        if name == 'currency' and hasattr(formats,'currencyFormat'):
            typeFormats[name]  = formats.currencyFormat
        elif name == 'date' and hasattr(formats,'dateFormat'):
            typeFormats[name]  = formats.dateFormat
        elif name == 'formula' and hasattr(formats,'formulaFormat'):
            typeFormats[name]  = formats.formulaFormat
        elif name == 'number' and hasattr(formats,'numberFormat'):
            typeFormats[name]  = formats.numberFormat
        elif name == 'percent' and hasattr(formats,'percentFormat'):
            typeFormats[name]  = formats.percentFormat
        elif name == 'text' and hasattr(formats, 'textFormat'):
            typeFormats[name] = formats.textFormat
        elif name == 'timestamp' and hasattr(formats,'timestampFormat'):
            typeFormats[name]  = formats.timestampFormat
        elif name == 'url' and hasattr(formats,'textFormat'):
            typeFormats[name]  = formats.textFormat
        else:
            print("What the what? -->",name)

    return typeFormats

def loadDataLabels(filename,formats):

    dataSet = []
    lookupReader = csv.reader(open(filename, newline=''), delimiter=',', quotechar='"')

    i = 0
    for row in lookupReader:
        i = i + 1
        if len(row) != 3:
            print("Error with row(",i,") :",row)
            continue
        myType = row[2]
        myFormat = formats.get(myType)
        myField = IEXField(row[0],row[1],myType,myFormat)
        dataSet.append(myField)

    # for d in dataSet:
    #   print(d)

    return dataSet