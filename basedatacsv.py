
import csv
import json

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

	def dividends( self ):
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

	def __init__(self, symbol, entry ) :

		if symbol == None or entry == None:
			print ('missing data')
			return

		self.symbol=symbol
		self.name= entry[2] # everyone has a name
		self.accounts = dict()
		self.addAccount( entry )

	def __str__(self) :
		return 'Ticker:' + self.symbol + ' ' + self.name + ' ' + str(self.numShares())

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

	def dividends( self ):
		total = 0.00
		for key, acct in self.accounts.items():
			total = total + acct.dividends()

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

		thisRow['Dividends Received'] = t.dividends()

		total_cost = t.cost()
		thisRow['Total Cost'] = total_cost;
		
		thisRow['Total Sold'] = t.sold()

		thisRow['Average Price'] = round( (abs(total_cost) / total_shares), 3 )
		print(str(thisRow))



if __name__ == "__main__":

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
			t = Ticker(s,row)
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
	printSymbols( symbols) 
	createSheet( symbols , unique_accounts )
