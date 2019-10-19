
import sys
import csv
import xlsxwriter
import json

class XLSFormats:

    def __init__(self, workbook = None, num = 2 ):
        self.workbook = workbook

        if num > 5:
            self.numberShades = 5
        else:
            self.numberShades = num

        self.dates = []
        self.texts = []
        self.numbers = []
        self.currency = []
        self.formulas = []
        self.percents = []

        # self.colors = ['#FFFFFF','#EEEEEE','#DDDDDD','#CCCCCC','#BBBBBB','#AAAAAA']
        self.colors = ['#D9E2F3','#FFFFFF'] #

        if workbook != None:
            self.header_format = workbook.add_format()
            self.header_format.set_font_size(14)
            self.header_format.set_bold()
            # self.header_format.set_bg_color('gray')
            self.header_format.set_bg_color('#4674C1')
            self.header_format.set_align('center')
            self.header_format.set_align('vcenter')

            for i in range(self.numberShades):

                date_format = workbook.add_format()
                date_format.set_font_size(14)
                date_format.set_num_format('mm/dd/yyyy')
                date_format.set_bg_color(self.colors[i])
                self.dates.append(date_format)

                text_format = workbook.add_format()
                text_format.set_font_size(14)
                text_format.set_bg_color(self.colors[i])
                self.texts.append(text_format)

                number_format = workbook.add_format()
                number_format.set_font_size(14)
                number_format.set_num_format(4)
                number_format.set_bg_color(self.colors[i])
                self.numbers.append(number_format)

                currency_format = workbook.add_format()
                currency_format.set_font_size(14)
                currency_format.set_num_format(8)
                currency_format.set_bg_color(self.colors[i])
                self.currency.append(currency_format)

                formula_format = workbook.add_format()
                formula_format.set_font_size(14)
                formula_format.set_bg_color(self.colors[i])
                self.formulas.append(formula_format)

                percent_fmt = workbook.add_format({'num_format': '0.00%'})
                percent_fmt.set_font_size(14)
                percent_fmt.set_bg_color(self.colors[i])
                self.percents.append(percent_fmt)

        else:
            self.header_format = None

    def dateFormat(self,row = 1):
        i = row % self.numberShades
        return self.dates[i]

    def textFormat(self,row = 1):
        i = row % self.numberShades
        return self.texts[i]

    def numberFormat(self,row = 1):
        i = row % self.numberShades
        return self.numbers[i]

    def currencyFormat(self, row=1):
        i = row % self.numberShades
        return self.currency[i]

    def formulaFormat(self, row=1):
        i = row % self.numberShades
        return self.formulas[i]

    def percentFormat(self, row=1):
        i = row % self.numberShades
        return self.percents[i]

    def headerFormat(self):
        return self.header_format

