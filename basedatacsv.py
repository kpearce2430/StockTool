
import os
import csv
import json
import urllib3
from datetime import date
# import transaction

# * * * * * * * * * * * * * * * *
class Lot:
	def __init__( self, numSharesToSell = 0.00, pricePerShare=0.00, soldDate = None ):	
		self.thisLot = dict()
		self.thisLot['numShares'] = numSharesToSell
		self.thisLot['pricePerShare'] = pricePerShare
		# thisLot['proceeds'] = numSharesToSell * pricePerShare
	
	def __str__(self):
		return json.dumps(self.thisLot)

	def proceeds(self):
		return float(self.thisLot['numShares'] * self.thisLot['pricePerShare'])
	
class Entry:
	
	def __init__( self, date = None, type = None, security = "", symbol = None, desc = None, shares = None, amount = None, account = None ):
	
		theEntry = dict()
		theEntry['entryDate'] = date # 0
		theEntry['entryType'] = type # 1
		theEntry['entrySecurity'] = security #2
		theEntry['entrySymbol'] = symbol # 3
		theEntry['entryDescription'] = desc # 4
		theEntry['entryShares'] = shares # 5
		theEntry['entryAmount'] = amount # 6
		theEntry['entryAmount'] = amount # 6
		theEntry['entryAccount'] = account #7
		theEntry['entryPricePerShare'] = 0.00
		theEntry['entryRemainingShares'] = 0.00

		if type == 'Buy' or type == 'Add Shares' or type == 'Reinvest Dividend':
			numShares = float(shares.replace(',',''))
			#
			# print(type, " ", numShares," of ",symbol," for ", account)
			if amount.isnumeric():
				myAmount = abs(float(amount))
			else:
				myAmount = 0.00
			# print("amount:", entry[6]," ",amount)
			theEntry['entryRemainingShares'] = numShares

			if numShares > 0.00:
				pps =  myAmount / numShares
				# print("pps:",pps," = ", amount , " / " , numShares )
				theEntry['entryPricePerShare'] = pps

		self.entry = theEntry
		self.soldLots = []

		# print('The Entry:' + str( self.entry))

	def printEntry(self):
		print(str(self))

	def Field(self,field):
		# if hasattr(self.entry,field):
			return str(self.entry[field])
		# else:
		#	return "Missing"

	def __str__(self) :
		return 'Entry:' + str(self.entry) 
	
	def description(self):
		# print(self.entry['entryDescription'])
		return self.entry['entryDescription']
	
	def security(self):
		return self.entry['entrySecurity']
	
	def symbol(self):
		return self.entry['entrySymbol']
		
	def price_per_share(self):
		# this is a befuddled way of calculating since quicken puts it in the
		# description
		myDescription = self.description()
		myWordList = myDescription.split(' ')
		if len(myWordList) != 4:
			print('Error - Unexpected Result from description:' + myDescription + ' Len[' + str(len(myWordList)) + ']')
			return self.entry['entryPricePerShare']
			# return 0.00
		
		# print(str(myWordList))
		ppsStr = myWordList[3]
		ppsStr.lstrip()
		# print('[' + ppsStr + '][' + str(float(ppsStr)) + ']')
		
		return float(ppsStr.replace(',',''))
		
	def shares(self):
		myShares = self.entry.get('entryShares')
		# print('Shares ' + self.type() + ' [' + myShares + ']')
		if myShares == "" or str(myShares) == 'None':
			return 0.0000
		else:
			return float(myShares.replace(',',''))

	def remainingShares(self):
		return self.numShares()
	
	def numShares(self):
		myShares = float(self.entry.get('entryRemainingShares'))

		# print("numShares:",myShares)
		if (myShares > 0.00):
			return myShares
		else:
			return 0.00

	def type(self):
		return self.entry.get('entryType')
	
	def dividendPaid(self):
		type =  self.type()
		amt = 0.00
		
		if  type == 'Dividend Income' or type == 'Reinvest Dividend' or type == 'Return of Capital':
			amt = self.amount()	
		
		return amt

	def amount(self):
		sAmount = self.entry.get('entryAmount')
		# print('Entry Amount:' + sAmount)
		try:
			return float( sAmount.replace(',',''))
		except:
			return 0.00

	def cost( self ) :
		amt = 0.00
		type =  self.entry.get('entryType')

		if  type == 'Buy' and self.remainingShares() > 0:
			#
			amt = self.remainingShares() * self.price_per_share()
			# amt = self.amount()
		
		return amt 
	
	def sellShares(self,numSharesToSell = 0.00, pricePerShare=0.00):
		totalShares = self.remainingShares()
		if  totalShares >= numSharesToSell:
			# if there are 50 shares to sell 100 remaining shares, remove 50 and return 0 
			# for the number of shares remaining to sell.
			# partial or full sale
			remainingShares = totalShares - numSharesToSell
			thisLot = Lot(numSharesToSell, pricePerShare)
			self.soldLots.append(thisLot)
			self.entry['entryRemainingShares'] = str(remainingShares)
			return 0.00 # no more shares remaining from sell
		else:
			# there are 50 shares remaining and 100 to sell, remove the 50 and
			# return there are 50 more to sell.
			remainingShares = numSharesToSell - totalShares
			thisLot = Lot(totalShares, pricePerShare)
			self.soldLots.append(thisLot)
			self.entry['entryRemainingShares'] = str(0.00)
			return round(remainingShares,4)
		
	def splitShares(self, oldShares=1.00, newShares=1.00, symbol=None):
		if self.remainingShares() <= 0:
			return # no shares remaining
		
		newRemainingShares =  ( float(self.remainingShares()) / oldShares ) * newShares; 
		self.entry['entryRemainingShares'] = str(newRemainingShares)
		# print('New Remaining Shares[' + self.entry['entryRemainingShares'] + ']')

	def entryDate(self):
		date_elem = self.entry['entryDate'].split('/')
		try:
			print("Date:",date_elem)
			if len(date_elem[2]) < 4:
				year = int(date_elem[2])
				# Since Quicken on exports 2 digit year, I took this simple approach
				if year > 50:
					year = 1900 + year
				else:
					year = 2000 + year
			else:
				year = int(date_elem[2])

			myDate = date(year,int(date_elem[0]),int(date_elem[1]))
			print("My Date:",myDate)
		except:
			print("Could not convert date:",date_elem)
			raise("Could not convert entry date ")

		return myDate


class Account:

	def __init__(self, name=None, entry=None ):


		if name == None or entry == None:
			print ('missing data')
			return

		print("Creating Account ", name)
		self.name = name
		self.entries = [] # dict()
		self.pending = []
		
		if isinstance(entry, Entry):
			self.addEntry(entry)
		else:
			raise("Not and entry")

		self.current_shares = 0
		

	def __str__(self) :
		return 'Account:' + self.name 

	def Name(self):
		return self.name

	def printAccount( self ):
		print( self )
		self.printEntries()


	def printEntries( self ):
		i = 0
		for e in self.entries:
			print( 'Entry:' +  str(i) + ':' + str(e))
			i = i + 1
			for lot in e.soldLots:
				print('Lot:' + str(lot) + 'p:' + str(lot.proceeds()))


	def addEntry( self, myEntry  ):
		# print(myEntry.type(), " For ", myEntry.symbol(), " Shares ", myEntry.shares())
		if isinstance( myEntry, Entry ):
			if myEntry.type() == 'Sell' or myEntry.type() == 'Short Sell':
				# print('Got a sell:' + str(myEntry))
				sellShares = abs(myEntry.shares())
				curentShares = self.numShares()
				if round(curentShares,4) >= round(sellShares,4) :
					
					# print ('Selling ' + str(sellShares) + ' Shares')
					pps = myEntry.price_per_share()
				
					for e in self.entries:
						if e.type() == 'Buy' or e.type() == 'Add Shares':
							sellShares = e.sellShares(sellShares,pps)
							# print ('Remaining ' + str(sellShares))
							if sellShares <= 0.00:
								break # for

					if sellShares > 0.0002:
						print('WARNING SHARES [' + myEntry.type() + '] WITHOUT Buy ' + str(sellShares) + ' ' + myEntry.symbol())
						print('Entry--> ' + str(myEntry))
						return False				
			
				else:
					# queue this entry until we have the shares
					# print('Queuing Entry[' + str(myEntry) + ']' + 'Sell: ' + str(sellShares))
					self.pending.append(myEntry)
					
			elif myEntry.type() == 'Stock Split':
				# 
				wordList = myEntry.description().split()
				# print(str(wordList))
				newShares = wordList[0]
				oldShares = wordList[2]
				for e in self.entries:
					e.splitShares(float(oldShares),float(newShares))
					
			elif myEntry.type() == 'Remove Shares':
				removeShares = abs(myEntry.shares())
				#
				# print("Removing Shares:", removeShares)
				currentShares = self.numShares()
				# print("Current Shares:",currentShares)
				# if currentShares == 0:
				#	self.printAccount()

				if round(currentShares,4) >= round(removeShares,4):
					# we can go ahead.  
					for e in self.entries:
						removeShares = e.sellShares(removeShares)
						if removeShares <= 0:
							break # for
						
					if removeShares > 0.00:
						print('WARNING SHARES [' + myEntry.type() + '] WITHOUT Buy ' + str(removeShares) + ' ' + myEntry.symbol())
						return False
									
				else: 
					# queue this entry until we have the shares
					print('Queuing Entry[' + str(myEntry) + ']')
					self.pending.append(myEntry)
				
			else:
				# print("Add:", str(myEntry))
				self.entries.append(myEntry)
				# self.printAccount()
				# print('Current ' + myEntry.symbol() + ',' + myEntry.type() + ',' + str(myEntry.shares()) + ' shares:' + str(self.numShares()) + ' pend: ' + str(self.totalPending()))


		else:
			raise( 'ERROR: entry is not an Entry' )
			return False


		if  (myEntry.type() == 'Buy' or myEntry.type() == 'Add Shares') and self.numShares() > self.totalPending():
			self.clearPending()

		return True

	def totalPending(self):
		total = 0.00
		for p in self.pending:
			if isinstance(p, Entry) == False:
				raise("Invalid Entry in Pending")

			amount = p.Field("entryAmount")
			total = total + abs(float(amount.replace(',','')))

		# print('Pending: ' + str(total))
		return total

	def clearPending(self):
		# print("Clear pending ",len(self.pending)," for ", self.Name())
		# self.printEntries()
		while len(self.pending) > 0:
			e = self.pending[0]
			# print(str(e))
			if self.addEntry(e):
				self.pending.remove(e)
				# print('Current ' + e[3] + ' shares:' + str(self.numShares()))
			else:
				raise ("ERROR Processing Pending Entry" + str(e))
				return
				
			
	def numShares( self ):
		total = 0.00
		for e in self.entries:	
			shares = e.numShares()
			total = total + shares

		return round(total,4)

	def dividends_paid( self ):
		total = 0.00
		for e in self.entries:
			total = total + e.dividendPaid()

		return round(total,2)

	def cost( self ) :
		total = 0.00
		for e in self.entries:
			type = e.type()
			if  type == 'Buy':
				if e.numShares() > 0:
					# print('Account ' + self.name + ' Entry Cost ' + str(e.amount()))
					total = total + e.cost()

		# print('Total Account Cost:' + str(total))
		return round(total,2)

	def sold( self ) :
		total = 0.00
		for e in self.entries:
			type = e.type()
			if  type == 'Sell' or type == 'Short Sell':
				# print ('Acct Sold:' + str(e.amount()))
				total = total + e.amount()

		return round(total,2)

	def firstBought(self):
		theDate = date.today()
		# print(theDate)
		for e in self.entries:
			if e.numShares() > 0:
				myDate = e.entryDate()
				if myDate < theDate:
					print("Account Setting date",myDate)
					theDate = myDate

		return theDate

# * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *  
class Ticker:

	def __init__(self, entry, http=None ) :


		if isinstance(entry,Entry) == False:
			raise('Invalid entry passed into Ticker')
			return

		self.symbol=entry.Field("entrySymbol")
		self.name= entry.Field("entrySecurity") # everyone has a name
		self.accounts = dict()
		self.http = http
		self.addToAccount( entry )
		self.token = os.environ.get("TOKEN")
		print(self.token)

	def worksheetName(self):
		# assert isinstance(self.name, s)
		return self.symbol + ' ' + self.name

	def __str__(self) :
		return 'Ticker:' + self.symbol + ' ' + self.name + ' ' + str(self.numShares()) + ' ' + str(self.current_dividend()) + ' ' + str(self.dividend_next12mo()) + ' ' + str(self.latest_price())

	def addToAccount(self, entry ):
		if isinstance(entry,Entry) == False:
			raise("Ticker.addToAccount did not get an Entry")
			# print("Ticker.addToAccount did not get an Entry:",entry)
			return

		acct_name = entry.Field("entryAccount")
		a = self.accounts.get(acct_name) 
		if a == None:
			# print( "added acct name:" + acct_name )
			a = Account(acct_name, entry )
			self.accounts[acct_name] = a
		else:
			# add the row
			# print( "just adding row:" )
			a.addEntry( entry )
		
			
		# if  a.numShares() > a.totalPending():
		#	self.clearPending()

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

	def clearPending(self):
		for key, acct in self.accounts.items():
			if acct.clearPending() != True:
				return False
			
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
			elif req_type == 'stats_data':
				return self.stats_data
			elif req_type == 'news':
				return self.news_data
			elif req_type == 'chart':
				return self.chart_data
			elif req_type == 'dividend_data':
				return self.dividend_data
			elif req_type == 'close':
				return self.close_data
			elif req_type == 'earnings':
				return self.earnings
			else:
				print('ERROR Invalid Type:' + req_type)
			return None
		
		# no attribute found, let's get the data.
		elif len(self.symbol) < 5:
			#
			# not a symbol we can get a quote on.
			# url = 'https://api.iextrading.com/1.0/stock/' + self.symbol.lower() + '/batch?types=quote,stats,earnings,news,chart,dividends,close&range=1y&last=3'
			# all the data
			#  TODO: need to build cache for each item below to save transaction costs.
			#
			url = 'https://cloud.iexapis.com/v1/stock/' + self.symbol.lower() + '/batch?types=quote,stats,dividends&range=1y&last=3&token=' + self.token # '
			#
			#
			# url = 'https://api.iextrading.com/1.0/stock/' + self.symbol.lower() + '/quote'
			r = self.http.request('GET',url)
			
			if r.status == 200:
				myData = json.loads(r.data.decode('utf-8'))
				if isinstance(myData, dict):
					self.quote_data =  myData.get('quote')
					self.stats_data =  myData.get('stats')
					self.news_data = myData.get('news')
					self.chart_data = myData.get('chart')
					self.dividend_data = myData.get('dividends')
					self.close_data = myData.get('close')
					self.earnings = myData.get('earnings')
					return self.get_data(req_type)
				else:
					print('Not JSON Format')
			else:
				print('Request Failed[' + str(r.status) + '] ' + str(r) )
				print("URL: ", url)
				
		self.quote_data =  None
		self.stats_data =  None
		self.news_data = None
		self.chart_data = None
		self.dividend_data = None
		self.close_data = None
		return None

	def get_earnings(self):
		return self.get_data('earnings')

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

	def lastest_eps(self):
		earnings = self.get_data('stats_data')
		# print("Here:", self.name,":",earnings)

		if earnings == None:
			print(self.name," No earnings")
			return None

		myearnings = earnings.get('ttmEPS')
		print(self.name,' earnings:',myearnings)
		if myearnings != None:
			return myearnings
		else:
			print("Missing LastestEPS:", self.name, ":", earnings)
			return None

	def firstBought(self):
		theDate = date.today()
		for key, acct in self.accounts.items():
			myDate = acct.firstBought()
			print("Account FB:",myDate)
			if myDate < theDate:
				theDate = myDate
				print("ticker setting date",theDate)

		return theDate

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
def createSheet( symbols, acct_list, details ):

	# allRows = []

	# header = 'Name,Symbol,Total Shares'
	# for a in acct_list:
	# 	header = header + ',' + a

	 # print(header)

	for key, value in symbols.items():
		
		# print('Symbol: ' + key)
		t  = symbols[key]
		
		# Since we can't guarantee the order of how the data is read, there is
		# a case where there can be pending sells or remove shares.  Take care
		# of those now.
		t.clearPending()
		
		total_shares = t.numShares()
		thisRow = dict()
		if total_shares < .001 :
			continue

		thisRow['Name'] = t.name
		thisRow['Symbol'] = t.symbol
		for a in acct_list:
			v = t.numSharesAccount( a )
			thisRow[a] = v 

		thisRow['Total Shares'] = total_shares;

		thisRow['Latest Price'] = t.latest_price()

		thisRow['Total Value'] = t.totalValue()

		thisRow['Dividends Received'] = t.dividends_paid()

		thisRow['Total Cost'] = abs(t.cost());
		
		# thisRow['Total Sold'] = t.sold()

		thisRow['Average Price'] = round( (abs(t.cost()) / total_shares), 3 )

		thisRow['Current Dividend'] = t.current_dividend()
		
		thisRow['Yearly Dividend'] = t.dividend_next12mo()
		
		thisRow['Net'] = t.net()

		thisRow['Latest EPS'] = t.lastest_eps()

		theDate = t.firstBought()
		thisRow['First Purchase'] = theDate.strftime('%m/%d/%Y')
		print("First Purchased:",theDate.strftime("%m/%d/%Y"))
		today = date.today()
		dateDelta = today - theDate
		thisRow['Days Owned'] = dateDelta.days

		# print(str(thisRow))
		
		details.append(thisRow)

def ProcessInfo( info, symbols, accounts, http):

	if isinstance(info, dict):
		s = info.get("security")

		# if s != 'CDC': # and s != 'UNP':
		#	return

		if s == str(None) or s == "Missing" or s == "DEAD" or s == 'Symbol':
			return

def ProcessEntry( entry, symbols, accounts, http):

	# print("ProcessEntry:", str(entry))
	if isinstance(entry, Entry):
		s = entry.Field("entrySymbol")
		# print("symbol:",s)

		# TODO:  Add a way to limit which symbols to look at.
		# if s != "PM":
		# 	return

		if s == str(None) or s == "Missing" or s == "DEAD" or s == 'Symbol':
			return

		t = symbols.get(s)
		# print("Symbol ",s," Ticker:", str(t))

		if t == None:
			#
			# create the ticker, add the account, add the row

			t	 = Ticker( entry, http)
			symbols[s] = t
			# print ("added:" + s )
		else:
			# print ("exists:" + s )
			t.addToAccount(entry)

		a = entry.Field("entryAccount")
		try:
			accounts.index(a)
		except ValueError:
			# print( 'adding account: ' + a )
			accounts.append(a)

	else:
		raise("Invalid parameter:",entry)

def ProcessRow(row, symbols, accounts, http):

	if isinstance(row, list):
		e = Entry(row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7])
		s = e.Field("entrySymbol")
		if s != 'PM': # and s != 'UNP':
			return
		ProcessEntry(e, symbols, accounts, http)

	else:
		raise("Invalid parameter:",row)


if __name__ == "__main__":


	http = urllib3.PoolManager()
	urllib3.disable_warnings()

	symbols = dict()
	unique_accounts = []
	details = []
	lookup = dict()

	lookupReader = csv.reader(open("lookup.csv", newline=''), delimiter=',', quotechar='"')
	for row in lookupReader:
		if len(row) == 2:
			# print(row[0],row[1])
			lookup[row[0]] = row[1]
		else:
			print("huh:", row)


	stockReader = csv.reader(open('quicken_data.csv', newline=''), delimiter=',',quotechar='"')

	i = 0
	for row in stockReader:

		# print(len(row),row[3])

		if (len(row) < 10):
			print("skipping[", i, "] items[", len(row), "] [", row, "]")
			continue

		# The first two columns from Quicken 18 are junk.
		# Get rid of them.
		del row[0]
		del row[0]

		k = row[3]
		s = lookup.get(k)
		row[3] = s

		ProcessRow(row,symbols,unique_accounts,http)

	unique_accounts.sort()

	for a in unique_accounts:
		print(a)

	printSymbols( symbols )
	
	createSheet(  symbols , unique_accounts, details )

	print( details )


