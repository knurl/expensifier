#!/usr/bin/python
import sys, getopt, csv, xlrd, xlwt, locale, os
from xlutils.copy import copy
from datetime import datetime, date

# A couple of constants. Should probably parameterise these at some point.
BUSINESSPURPOSE = "Sales"
DEPARTMENT = 'Sales'

# Styles for numbers, text written into the spreadsheet
currencyStyle = xlwt.easyxf("font: color black, bold 1", "#,###.00")
textStyle = xlwt.easyxf("font: color blue")

class InvalidVersion(Exception):
	def __init__(self, value):
		self.value = value
	def __str__(self):
		return repr(self.value)

class OverflowException(Exception):
	def __init__(self, value):
		self.value = value
	def __str__(self):
		return repr(self.value)

class InvalidCSVCurrency(Exception):
	def __init__(self, value):
		self.value = value
	def __str__(self):
		return repr(self.value)

# base class: Expenses always have to have dates, merchant, etc
class Expense:
	def __init__(self, expenseType, date, description, amount, merchant, origCurrency, origAmount):
		self.expenseType = expenseType
		self.date = date
		self.description = description
		self.amount = float(amount)
		self.merchant = merchant
		self.origCurrency = origCurrency
		self.origAmount = float(origAmount)

	def __repr__(self):
		return "(type: " + self.expenseType + "; merchant: " + \
			self.merchant + "; date: " + \
			str(self.date) + "; description: \"" + \
			self.description + "\"; amount: " + \
			str(self.amount) + "; origCurrency: " + \
			self.origCurrency + "; origAmount: " + \
			self.origAmount + ")"

def char_range(c1, c2):
	"""Generates the characters from `c1` to `c2`, inclusive."""
	for c in xrange(ord(c1), ord(c2)+1):
		yield chr(c)

# hide all the details of the current expense report format here so
# it can be changed more easily if the format changes in the future
class ExpenseV3:
	v3SheetName = "Expense Report Revised V3"
	mapExpenseTypeToColumnLetter = { \
		'breakfast':'C', \
		'lunch':'D', \
		'dinner':'E', \
		'hotel':'G', \
		'air':'H', \
		'rail':'I', \
		'carRental':'J', \
		'taxi':'K', \
		'parkingToll':'L', \
		'phone':'M', \
		'other':'N' \
	}

	mapMandatoryFillToAddr = { \
			'businessPurposeAddr':'H1', \
			'nameAddr':'L1', \
			'departmentAddr':'N2', \
			'periodCoveredAddr':'N3'
	}

	travelExpensesRowStart = 6
	travelExpensesMax = 8
	entertainmentExpensesRowStart = 18
	entertainmentExpensesMax = 8
	miscellaneousExpensesRowStart = 30
	miscellaneousExpensesMax = 11

	# writer must be a function that works on the already-provided sheet st
	def recreateFormulas(self, writer):
		for i in range(6, 14):
			writer(self.st, 'P' + str(i), 'SUM(G%d:O%d)' % (i, i))
			writer(self.st, 'F' + str(i), 'SUM(C%d:E%d)' % (i, i))
		for i in char_range('F', 'P'):
			writer(self.st, i + '14', 'SUM(%s6:%s13)' % (i, i))
		writer(self.st, 'P26', 'SUM(P18:P25)')
		for col in 'H', 'L', 'O':
			writer(self.st, col + '41', 'SUM(%s30:%s40)' % (col, col))
		for i in range(30, 35):
			writer(self.st, 'L' + str(i), chr(ord('F')+i-30) + '14')
		writer(self.st, 'L35', 'K14+L14')
		writer(self.st, 'L36', 'M14')
		writer(self.st, 'L37', 'N14+O14')
		writer(self.st, 'L38', 'P26')
		writer(self.st, 'L39', 'P52')
		writer(self.st, 'L40', 'H41')
		for i in range(30, 41):
			writer(self.st, 'O' + str(i), 'L' + str(i) + '*N' + str(i))
		writer(self.st, 'C44', 'L41')

	def writeTravelExp(self, writer, rowIndex, date, desc, expenseMap):
		row = self.travelExpensesRowStart + rowIndex
		writer(self.st, 'A' + str(row), str(date))
		writer(self.st, 'B' + str(row), str(desc))
		for category, column in self.mapExpenseTypeToColumnLetter.iteritems():
			if category in expenseMap:
				writer(self.st, column + str(row), \
				expenseMap[category], \
				currencyStyle)

	def writeEntertainmentExp(self, writer, rowIndex, date, desc, amount, merchant):
		row = self.entertainmentExpensesRowStart + rowIndex
		writer(self.st, 'A' + str(row), str(date))
		writer(self.st, 'B' + str(row), BUSINESSPURPOSE)
		writer(self.st, 'D' + str(row), str(desc))
		writer(self.st, 'J' + str(row), str(merchant))
		writer(self.st, 'P' + str(row), amount, currencyStyle)
		
	def writeMiscellaneousExp(self, writer, rowIndex, date, desc, amount, merchant):
		row = self.miscellaneousExpensesRowStart + rowIndex
		writer(self.st, 'A' + str(row), str(date))
		writer(self.st, 'B' + str(row), str(merchant) + ': ' + str(desc))
		writer(self.st, 'F' + str(row), DEPARTMENT)
		writer(self.st, 'H' + str(row), amount, currencyStyle)

	def writeMandatoryData(self, writer, currency, name, lowDate, highDate):
		writer(self.st, 'O29', currency)
		writer(self.st, 'B44', 'TOTAL EXPENSE ' + currency)
		writer(self.st, 'E44', 'Not Applicable - Reimbursement in ' + currency)
		writer(self.st, 'G44', 'Not Applicable - Reimbursement in ' + currency)
		writer(self.st, 'H1', BUSINESSPURPOSE)
		writer(self.st, 'M1', name)
		writer(self.st, 'N2', DEPARTMENT)
		if lowDate != None and highDate != None:
			writer(self.st, 'N3', str(lowDate) + ' to ' + str(highDate))
		writer(self.st, 'E48', str(date.today()))

	def __sanityCheck(self):
		# should fill this in at some point!
		pass

	def getExpenseType(self, expensifyField):
		return self.mapExpensifyFieldToExpenseType(expensifyField)

	# We want to create an expense accumulator, that keeps entries based on date and
	# limits the number of expenses allowed
	class AccumulatedDailyExpenseSet:
		class DailyExpenseAccumulator:
			def __init__(self, date, description, expenseMap):
				self.date = date
				self.description = description
				self.expenseMap = expenseMap

			def __repr__(self):
				return "<<" + str(self.date) + "," + \
					self.description + "," + \
					str(self.expenseMap) + ">>"

		def __init__(self, maxEntries):
			self.maxEntries = maxEntries
			self.accumulatorSet = {}

		def __formDescription(self, merchant, description):
			d = description
			if len(d) > 0:
				return merchant + ': ' + d
			return d

		def __len__(self):
			return len(self.accumulatorSet)

		def combine(self, expense):
			if expense.date in self.accumulatorSet:
				# first get the expense
				accum = self.accumulatorSet[expense.date]

				# add the description
				if len(expense.description) > 0:
					if len(accum.description) > 0:
						accum.description += ' / '
					accum.description += \
						self.__formDescription(expense.merchant, \
						expense.description)

				# now add in the amount to the correct keyed entry
				expenseMap = accum.expenseMap
				if expense.expenseType in expenseMap:
					expenseMap[expense.expenseType] += expense.amount
				else:
					expenseMap[expense.expenseType] = expense.amount
			else:
				if len(self.accumulatorSet) >= self.maxEntries:
					raise OverflowException("too many dates for expense accumulator")
				self.accumulatorSet[expense.date] = \
					self.DailyExpenseAccumulator(expense.date, \
					self.__formDescription(expense.merchant, \
					expense.description), \
					dict([(expense.expenseType, expense.amount)]))

		def getAccumulators(self):
			return [v for (k, v) in \
				sorted(self.accumulatorSet.items())]

		def __repr__(self):
			s = "{" + str(len(self.accumulatorSet)) + " entries{"
			v = self.accumulatorSet.values()
			for entry in v:
				s += str(entry)
			s += "}}"
			return s


	# We want a collection representing a fixed set of expenses, where there can be
	# multiple entries for a particular date (date is not a key, in other words)
	class FixedExpenseSet:
		def __init__(self, maxEntries):
			self.maxEntries = maxEntries
			self.expenseSet = []

		def combine(self, expense):
			if len(self.expenseSet) >= self.maxEntries:
				raise OverflowException("too many dates for fixed expense set")
			self.expenseSet.append(expense)

		def getValues(self):
			return self.expenseSet

		def __len__(self):
			return len(self.expenseSet)

		def __repr__(self):
			s = "(size=%d(" % len(self.expenseSet)
			for entry in self.expenseSet:
				s += str(entry)
			s += "))"
			return s

	def isEmpty(self):
		return len(self.travelExp) == 0 and \
			len(self.entertainmentExp) == 0 and \
			len(self.miscellaneousExp) == 0

	def addCurrencyCost(self, exp):
		# if the credit card charges a currency conversion
		# charge (and most do), then add it in here.
		if currencyUpliftPerc > 0.0 and \
			homeCurrency != exp.origCurrency:
			currencyCostAdjustment = round(exp.amount * \
				currencyUpliftPerc, 2)
			self.currencyCost += currencyCostAdjustment
			print "Adding %s%.2f of currency cost to " \
				"amount %s%.2f (orig %s) " \
				"on %s" % \
				(homeCurrency, currencyCostAdjustment, \
				homeCurrency, exp.amount, exp.origCurrency,
				str(exp.date))

	def resetLowHighDates(self, exp):
		self.low, self.high = dateBounds(exp.date, self.low, self.high)

	# For all combiner functions below:
	# The currency cost needs to be put in the Other field of the
	# travel expense section. If we can use any date in that section
	# to store the currency costs, then do it. Otherwise, use a date
	# from the entertainment or miscellaneous sections (but still
	# stick it in the travel expenses section)

	def combineTravelExp(self, exp):
		self.travelExp.combine(exp)
		self.addCurrencyCost(exp)
		self.anyDate = exp.date
		self.resetLowHighDates(exp)

	def combineEntertainmentExp(self, exp):
		self.entertainmentExp.combine(exp)
		self.addCurrencyCost(exp)
		# ONLY if empty; always prefer a date from travel sec
		if self.anyDate == None:
			self.anyDate = exp.date
		self.resetLowHighDates(exp)

	def combineMiscellaneousExp(self, exp):
		self.miscellaneousExp.combine(exp)
		self.addCurrencyCost(exp)
		# ONLY if empty; always prefer a date from travel sec
		if self.anyDate == None:
			self.anyDate = exp.date
		self.resetLowHighDates(exp)

	def combine(self, exp):
		if exp.expenseType == 'entertainment':
			self.combineEntertainmentExp(exp)
		elif exp.expenseType == 'miscellaneous':
			self.combineMiscellaneousExp(exp)
		else:
			self.combineTravelExp(exp)

	def save(self):
		assert not self.isEmpty()

		# Insert the currency cost
		assert self.anyDate != None
		try:
			if self.currencyCost > 0.0:
				explanation = "%s%.2f in total currency conversion charges" % \
					(homeCurrency, self.currencyCost)
				currExp = Expense('other', self.anyDate, explanation, \
					self.currencyCost, "CC", homeCurrency, self.currencyCost)
				xpen.travelExp.combine(currExp)
				print explanation
		except OverflowException:
			# anyDate should be set to (i) any date in the accumulator set,
			# if it is nonempty, or otherwise (ii) any other date. So
			# the currency charge here should *not* be creating a new entry
			# in the accumulator set, and therefore should not trigger a
			# *new* overflow.
			print "Shouldn't get overflow exception here if anyDate is working!"
			print "anyDate is", str(self.anyDate), "; travelExp:",
			for exp in self.travelExp.getAccumulators():
				print exp.date,
			print "entertainmentExp: ",
			for exp in self.entertainmentExp.getValues():
				print exp.date,
			print "miscellaneousExp: ",
			for exp in self.miscellaneousExp.getValues():
				print exp.date,
			raise
		except:
			raise

		# Now write everything into the Excel spreadsheet
		rowIndex = 0
		for exp in self.travelExp.getAccumulators():
			self.writeTravelExp(writer, rowIndex, exp.date, \
					exp.description, exp.expenseMap)
			rowIndex += 1
		rowIndex = 0
		for exp in self.entertainmentExp.getValues():
			if len(exp.description) == 0:
				print "WARNING: Entertainment expenses MUST have descriptions"
			self.writeEntertainmentExp(writer, rowIndex, exp.date, \
				exp.description, exp.amount, exp.merchant)
			rowIndex += 1
		rowIndex = 0
		for exp in self.miscellaneousExp.getValues():
			if len(exp.description) == 0:
				print "WARNING: Miscellaneous expenses MUST have descriptions"
			self.writeMiscellaneousExp(writer, rowIndex, exp.date, \
				exp.description, exp.amount, exp.merchant)
			rowIndex += 1

		self.writeMandatoryData(writer, homeCurrency, yourName, \
			self.low, self.high)

		# Save the workbook
		print "Saving to " + self.fnOutputSpreadsheet
		self.wb.save(self.fnOutputSpreadsheet)

	def __init__(self, wb, st, fnOutputSpreadsheet):
		try:
			self.anyDate = None
			self.low = None
			self.high = None
			self.wb = wb
			self.st = st
			self.fnOutputSpreadsheet = fnOutputSpreadsheet
			self.currencyCost = 0.0

			self.travelExp = \
				self.AccumulatedDailyExpenseSet(self.travelExpensesMax)
			assert len(self.travelExp) == 0

			self.entertainmentExp = \
				self.FixedExpenseSet(self.entertainmentExpensesMax)
			assert len(self.entertainmentExp) == 0

			self.miscellaneousExp = \
				self.FixedExpenseSet(self.miscellaneousExpensesMax)
			assert len(self.miscellaneousExp) == 0
		except:
#			raise InvalidVersion("No V3 worksheet available")
			raise
		self.__sanityCheck()

# hide all the details of the Expensify expense report format here so
# it can be changed more easily if the format changes in the future
class ExpensifyFormatV1:
	mapExpensifyFieldToExpenseType = { \
			'Entertainment':'entertainment', \
			'Lodging':'hotel', \
			'Meals - Breakfast':'breakfast', \
			'Meals - Dinner':'dinner', \
			'Meals - Lunch':'lunch', \
			'Miscellaneous (not travel related)':'miscellaneous', \
			'Other (travel related)':'other', \
			'Phone':'phone', \
			'Transport - Air':'air', \
			'Transport - Car Rental':'carRental', \
			'Transport - Fuel':'parkingToll', \
			'Transport - Parking':'parkingToll', \
			'Transport - Rail':'rail', \
			'Transport - Taxi':'taxi', \
			'Transport - Toll':'parkingToll' \
	}

	def initExpensifyDump(self, f):
		self.rdr = csv.DictReader(f)

	def convertExpense(self, exp):
		try:
			category = exp['Category']
			date = datetime.strptime(exp['Timestamp'], "%Y-%m-%d %H:%M:%S").date()
			description = exp['Comment']
			amount = locale.atof(exp['Amount']) # the locale conversion!
			merchant = exp['Merchant']
			origCurrency = exp['Original Currency']
			# We should really have an 'original locale' field, but
			# this is only used for an assertion later, so this will
			# suffice
			origAmount = locale.atof(exp['Original Amount'])
			try:
				expenseType = self.mapExpensifyFieldToExpenseType[category]
			except:
				raise InvalidVersion('Category for expense looks wrong: "%s"' \
					% category)

			return Expense(expenseType, date, description, \
					amount, merchant, origCurrency, \
					origAmount)
		except InvalidVersion:
			raise
		except Exception as ex:
			raise InvalidVersion('Cannot locate fields in Expensify dump:' + str(ex))

	def getExpenses(self):
		return map(self.convertExpense, self.rdr)

	def getExpensifyCategories(self):
		return self.mapExpensifyFieldToExpenseType.keys()

# should put some unit testing code here!

#########################
# XLWT HELPER FUNCTIONS #
#########################

# Excel column letter to xlwt equivalent
def colConvert(x):
	# only handle the simple case
	assert len(x) == 1
	assert 'A' <= x <= 'Z'
	return ord(x) - ord('A')

# Excel row number to xlwt equivalent
def rowConvert(x):
	assert x > 0
	return x - 1

# Excel alphanumeric column:row formats (e.g. "A12") to xlwt equivalent
# Only handling columns <= 26 for now for simplicity's sake, guarded
# by asserts
def addressConvert(address):
	col = address[0:1]
	assert col.isalpha()
	assert col.isupper()
	row = address[1:]
	assert row.isdigit()
	return (rowConvert(int(row)), colConvert(col))

# General writer function for writing into spreadsheet
# Handed into ExpenseV3 class so it doesn't need to
# know about the specifics of xlwt
def writer(st, address, value, style = textStyle):
	r, c = addressConvert(address)
	st.write(r, c, value, style)

# same as above, but for writing formulas
def formulaWriter(st, address, value, style = currencyStyle):
	writer(st, address, xlwt.Formula(value), style)

#########################
# THE PROCESSING BEGINS #
#########################

expensifyWrapper = ExpensifyFormatV1()

usage = sys.argv[0] + """ [-l <locale>] [-c <currency>] [-u <curr_uplift>] <name> <expensify_dump> <input_sheet>

locale:		Set locale associated with expensify_dump. Default is en_US.
currency:	3-letter currency string. Default is USD.
curr_uplift:	Floating point uplift % charged by credit card, e.g., 4.5
name:		Your name, e.g. "John Hancock"
expensify_dump:	Dumpfile from Expensify in .cvs format--must contain only ascii chars
input_sheet:	Original expense spreadsheet in .xls format

EXAMPLE: """ + sys.argv[0] + """ \"Fred Astaire\" ~/Downloads/Bulk_Export_id_DEFAULT_CSV.csv expense-form.xls

Expensify config: You need to do the following in Expensify to use this tool:
1. Go to Settings->Preferences and set default currency. This is what Expensify
   will convert foreign currency transactions to.
2. In the same panel, set Cash Conversion Surcharge to 0. If you want to set a
   currency conversion cost, you do it in the tool by using the -u option,
   because Billing wants the conversion cost broken out and put in the Other
   category and -u will do this for you automatically.
3. Make sure your categories are set to the following. Every character needs to
   be exactly as written below. If these are not set correctly, you'll get an
   exception.
""" + str(expensifyWrapper.getExpensifyCategories())

try:
	opts, args = getopt.getopt(sys.argv[1:],"hl:c:u:")
except getopt.GetoptError:
	print usage
	sys.exit(2)

# Defaults
expensifyLocale = 'en_US'
homeCurrency = 'USD'
currencyUpliftPerc = 0.0

for opt, arg in opts:
	if opt == '-h':
		print usage
		sys.exit()
	elif opt == "-l":
		expensifyLocale = arg
	elif opt == "-c":
		homeCurrency = arg
	elif opt == "-u":
		currencyUpliftPerc = float(arg) / 100.0

if len(args) < 3:
	print usage
	sys.exit(2)

yourName = args[0]
fnExpensifyDump = args[1]

(fnExpensifyDumpPath, fnExpensifyDumpExt) = os.path.splitext(fnExpensifyDump)
if fnExpensifyDumpExt != ".csv":
	print "expensify_dump file must end in .csv"
	print usage
	sys.exit(2)

fnBlankSpreadsheet = args[2]

fnOutputSpreadsheetStem = fnExpensifyDumpPath

# ESSENTIAL for interpreting currencies from Expensify. Expensify cannot 
# be configured to NOT print currencies out in a locale-specific format!
# Eventually I will need to make the following a variable, tied to some
# user-defined parameter
locale.setlocale(locale.LC_ALL, expensifyLocale)

# Helper for finding youngest, oldest dates amongst expenses
# Used for filling in date range onto sheet
def dateBounds(date, low, high):
	if low == None or date < low:
		low = date
	if high == None or date > high:
		high = date
	return low, high

def getExpenseSheetCopy(rb, fnBlankSpreadsheet, fnOutputSpreadsheet):
	wb = copy(rb) # copy from read-only spreadsheet to output form

	# get first sheet and verify it's the one we expect,
	# then instantiate sheet as a ExpenseV3 sheet
	wbst = wb.get_sheet(0)
	assert wbst.name == ExpenseV3.v3SheetName
	xpen = ExpenseV3(wb, wbst, fnOutputSpreadsheet)
	assert xpen.isEmpty()

	# We need to recreate all the formulas. Sigh.
	# xlutils doesn't have the ability to convert over formulas,
	# so we have to recreate all of them!
	xpen.recreateFormulas(formulaWriter)

	return xpen

with open(fnExpensifyDump, 'rb') as fExpensifyDump:
	outputSheetCounter = 0

	# pull in the expenses from the Expensify csv dump as an iterator
	expensifyWrapper.initExpensifyDump(fExpensifyDump)
	expenses = expensifyWrapper.getExpenses()

	# only once, get a read-only copy of the expense sheet from the filename
	# of the blank spreadsheet
	rb = xlrd.open_workbook(fnBlankSpreadsheet, formatting_info=True)

	# get an expense sheet copy
	outputSheetCounter += 1
	fnOutputSpreadsheet = fnOutputSpreadsheetStem + '-' + \
		str(outputSheetCounter) + ".xls"
	xpen = getExpenseSheetCopy(rb, fnBlankSpreadsheet, fnOutputSpreadsheet)

	# process each expense one by one
	for exp in expenses:
		# Either the target and source currencies are identical, OR
		# Expensify should have done a conversion--otherwise Expensify
		# was set up wrong!
		if homeCurrency == exp.origCurrency:
			if exp.amount != exp.origAmount:
				explanation = ("Check your default currency in "
						"Expensify, and compare to the "
						"currency you're specifying. "
						"Output currency %s and Expensify "
						"currency %s match, but amounts "
						"%.2f and %.2f are different" %
						(homeCurrency, exp.origCurrency,
						exp.amount, exp.origAmount))
				raise InvalidCSVCurrency(explanation)
		elif exp.amount == exp.origAmount:
				explanation = ("Check your default currency in "
						"Expensify, and compare to the "
						"currency you're specifying. "
						"Output currency %s and Expensify "
						"currency %s differ, but amounts "
						"%.2f and %.2f are the same" %
						(homeCurrency, exp.origCurrency,
						exp.amount, exp.origAmount))
				raise InvalidCSVCurrency(explanation)


		try:
			xpen.combine(exp)
		except OverflowException:
			# if we couldn't fit in the expense, start a new sheet
			assert not xpen.isEmpty()
			xpen.save()

			outputSheetCounter += 1
			fnOutputSpreadsheet = fnOutputSpreadsheetStem + '-' + \
				str(outputSheetCounter) + ".xls"
			xpen = getExpenseSheetCopy(rb, fnBlankSpreadsheet, \
				fnOutputSpreadsheet)
			# Don't forget to add that expense to the new sheet!
			assert xpen.isEmpty()
			xpen.combine(exp)
		except:
			raise

	assert not xpen.isEmpty()
	xpen.save()
