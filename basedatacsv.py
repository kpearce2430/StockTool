
import csv
import json
import os
import json
import urllib3
import time
from datetime import date
from pip._vendor.html5lib.filters.sanitizer import data_content_type

# * * * * * * * * * * * * * * * *

class Account:

	def __init__(self, name=None, entry=None ):

		if name == None or entry == None:
			print ('missing data')
			return
		self.name = name
		self.entries = [] # dict()
		self.entries.append(entry)
		# print ('account[' + name + '] added entry:' + str(entry) )

	def __str__(self) :
		return 'Account:' + self.name 

	def printAccount( self ):
		print( self )
		self.printEntries()

	def printEntries( self ):
		i = 0
		for e in self.entries:
			print( 'Entry:' +  str(i) + ':' + str(e))
			i = i + 1

	def addEntry( self, entry ):
		if isinstance( entry, list ):
			self.entries.append(entry)
			# print ('account[' + self.name + '] added entry:' + str(entry) )
			# self.printEntries()
		else:
			print( 'ERROR: entry is not a list' )

	def numShares( self ):
		total = 0.00
		for e in self.entries:	
			# print ( 'e5=' +  e[5] )
			if  str(e[5]) == 'None' or str(e[5]) == "" :
				continue
			shares = float( e[5].replace(',',''))
			total = total + shares
			# print ( total, shares )
	
		return round(total,4)

	def dividends_paid( self ):
		total = 0.00
		for e in self.entries:
			type = e[1]
			if  type == 'Dividend Income' or type == 'Reinvest Dividend' or type == 'Return of Capital':
				amt = float( e[6].replace(',',''))
				total = total + amt

		return round(total,2)

	def cost( self ) :
		total = 0.00
		for e in self.entries:
			type = e[1]
			if  type == 'Buy':
				amt = float( e[6].replace(',',''))
				total = total + amt

		return round(total,2)

	def sold( self ) :
		total = 0.00
		for e in self.entries:
			type = e[1]
			if  type == 'Sold' or type == 'Short Sell':
				amt = float( e[6].replace(',',''))
				total = total + amt

		return round(total,2)

# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *  
class Ticker:

	def __init__(self, symbol, entry, http=None ) :

		if symbol == None or entry == None:
			print ('missing data')
			return

		self.symbol=symbol
		self.name= entry[2] # everyone has a name
		self.accounts = dict()
		self.http = http
		self.addAccount( entry )

	def __str__(self) :
		return 'Ticker:' + self.symbol + ' ' + self.name + ' ' + str(self.numShares()) + ' ' + str(self.current_dividend()) + ' ' + str(self.dividend_next12mo()) + ' ' + str(self.quote())

	def addAccount(self, entry ):
		acct_name = row[7]
		a = self.accounts.get(acct_name) 
		if a == None:
			# print( "added acct name:" + acct_name )
			myAcct = Account(acct_name, entry )
			self.accounts[acct_name] = myAcct
		else:
			# add the row
			# print( "just adding row:" )
			a.addEntry( entry )

	def printAccounts( self ):
		for key, acct in self.accounts.items():
			# print( 'key:' +  key  + 'value:' + str(acct))
			acct.printAccount()

	def numSharesAccount( self, acct_name ):
		a = self.accounts.get( acct_name )

		if a != None:
			return a.numShares()
		else:
			return 0.00

	def numShares( self ):
		total = 0.00
		for key, acct in self.accounts.items():
			total = total + acct.numShares()

		return round(total,4)

	def dividends_paid( self ):
		total = 0.00
		for key, acct in self.accounts.items():
			total = total + acct.dividends_paid()

		return round(total,2)

	def cost( self ):
		total = 0.00
		for key, acct in self.accounts.items():
			total = total + acct.cost()

		return round(total,2)

	def sold( self ):
		total = 0.00
		for key, acct in self.accounts.items():
			total = total + acct.sold()

		return round(total,2)

	# stock_ticker = self.symbol.lower()
	# url = 'https://api.iextrading.com/1.0/stock/' + self.symbol.lower() + '/dividends/1y'
	# url = 'https://api.iextrading.com/1.0./stock/' + stock_ticker + '/dividends/1y'
	# print('url:' + url )
	
	def current_dividend( self ):
	
		if hasattr(self, 'dividend_amount'):
			return self.dividend_amount
		else:	
			myData = self.get_data('dividend_data')

			if isinstance(myData, list):
				# print('got my list')
				# self.dividend_data = myData
				self.dividend_multiplier = float(len(myData))
				firstOne = True
				for myJson in myData:
					# print(myJson)
					if firstOne == True:
						self.dividend_amount =  float(myJson.get('amount'))
						# print( 'amount:' + str(self.dividend_amount))
						# print( 'yearly:' + str(self.dividend_amount * self.dividend_multiplier))
						firstOne = False
						return self.dividend_amount
			
			elif myData == None:
				self.dividend_multiplier = 0.00
				self.dividend_amount = 0.00				
				return 0.00
			else:
				print('Unexpected Div Data[' + str(myData) + ']')
				
		self.dividend_multiplier = 0.00
		self.dividend_amount = 0.00
		return 0.00

	def dividend_next12mo(self):
		
		return float( self.current_dividend() * self.dividend_multiplier )
	
	def get_data(self,req_type):
		
		# print('getting[' + req_type + ']')
		if hasattr(self, req_type):
			if req_type == 'quote_data':
				return self.quote_data
			elif req_type == 'stats':
				return self.stats_data
			elif req_type == 'news':
				return self.news_data
			elif req_type == 'chart':
				return self.chart_data
			elif req_type == 'dividend_data':
				return self.dividend_data
			elif req_type == 'close':
				return self.close_data
			else:
				print('ERROR Invalid Type:' + req_type)
			return None
		
		# no attribute found, let's get the data.
		elif len(self.symbol) < 5:
			#
			# not a symbol we can get a quote on.
			url = 'https://api.iextrading.com/1.0/stock/' + self.symbol.lower() + '/batch?types=quote,stats,news,chart,dividends,close&range=1y&last=3'
			#
			# url = 'https://api.iextrading.com/1.0/stock/' + self.symbol.lower() + '/quote'
			r = http.request('GET',url)
			
			if r.status == 200:
				myData = json.loads(r.data.decode('utf-8'))
				if isinstance(myData, dict):
					self.quote_data =  myData.get('quote')
					self.stats_data =  myData.get('stats')
					self.news_data = myData.get('news')
					self.chart_data = myData.get('chart')
					self.dividend_data = myData.get('dividends')
					self.close_data = myData.get('close')
					return self.get_data(req_type)
				else:
					print('Not JSON Format')
			else:
				print('Request Failed')
				
		self.quote_data =  None
		self.stats_data =  None
		self.news_data = None
		self.chart_data = None
		self.dividend_data = None
		self.close_data = None
		return None
			
	def get_quote(self):	
		return self.get_data('quote_data')
		
		
	def latest_price(self):
		if hasattr(self, 'latestPrice') != True:

			quoteData = self.get_quote()
			if quoteData != None:
				self.latestPrice = float(quoteData.get('latestPrice'))
			else:
				self.latestPrice = 0.00
						
		return self.latestPrice
	
	def totalValue(self ):
		return round(float(self.latest_price()* float(self.numShares())),2)
	
	def net(self):
		return round(float(self.totalValue()) + float(self.cost()),2)
	
#				
# * * * * * * * * * * * 
#
def printSymbols( symbols ):

	print('Symbols:')
	for key, value in symbols.items():
		print( key + ':' + str(value))
		t  = symbols[key]
		t.printAccounts()

#
#
#
def createSheet( symbols, acct_list ):

	allRows = []

	header = 'Name,Symbol,Total Shares' 
	for a in acct_list:
		header = header + ',' + a

	print(header)

	for key, value in symbols.items():
		
		# print('Symbol: ' + key)
		t  = symbols[key]
		total_shares = t.numShares()
		thisRow = dict()
		if total_shares < .0001 :
			continue

		thisRow['Name'] = t.name
		thisRow['Symbol'] = t.symbol
		for a in acct_list:
			v = t.numSharesAccount( a )
			thisRow[a] = v 

		thisRow['Total Shares'] = total_shares;

		thisRow['Dividends Received'] = t.dividends_paid()

		total_cost = t.cost()
		thisRow['Total Cost'] = total_cost;
		
		thisRow['Total Sold'] = t.sold()

		thisRow['Average Price'] = round( (abs(total_cost) / total_shares), 3 )

		thisRow['Current Dividend'] = t.current_dividend()
		
		thisRow['Yearly Dividend'] = t.dividend_next12mo()
		
		thisRow['Latest Price'] = t.latest_price()
		
		thisRow['Total Value'] = t.totalValue()
		
		thisRow['Net'] = t.net()

		print(str(thisRow))
		
		allRows.append(thisRow)



if __name__ == "__main__":


	http = urllib3.PoolManager()
	urllib3.disable_warnings()

	symbols = dict()
	unique_accounts = []

	stockReader = csv.reader(open('quicken_data.csv', newline=''), delimiter=',',quotechar='"')

	for row in stockReader:
		# print(len(row),row[3])

		s = row[3]
		if s == str(None) or s == "Missing" or s == "DEAD" or s == 'Symbol' :
			continue

		t = symbols.get(s)
		if t == None:
			#
			# create the ticker, add the account, add the row
			t = Ticker(s,row, http)
			symbols[s] = t
			# print ("added:" + s )
		else:
			# print ("exists:" + s )
			t.addAccount(row)
		
		a = row[7]
		try:
			unique_accounts.index(a) 
		except ValueError:
			# print( 'adding account: ' + a )
			unique_accounts.append(a)
			# for i in unique_accounts:
			# 	print( i )

	unique_accounts.sort()
	
	# printSymbols( symbols) 
	
	createSheet(  symbols , unique_accounts )

	# text = commonRequestCall('https://api.iextrading.com/1.0/stock/psec/dividends/1y', disable_warnings=False,  myTimeout=15.00):
	# print(text)
