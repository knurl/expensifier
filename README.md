expensifier
===========

Convert CSV dumps from Expensify into an employer-friendly spreadsheet

./expensifier.py [-l <locale>] [-c <currency>] [-u <curr_uplift>] <name> <expensify_dump> <input_sheet>

locale:		Set locale associated with expensify_dump. Default is en_US.
currency:	3-letter currency string. Default is USD.
curr_uplift:	Floating point uplift % charged by credit card, e.g., 4.5
name:		Your name, e.g. "John Hancock"
expensify_dump:	Dumpfile from Expensify in .cvs format--must contain only ascii chars
input_sheet:	Original expense spreadsheet in .xls format

EXAMPLE: ./expensifier.py "Fred Astaire" ~/Downloads/Bulk_Export_id_DEFAULT_CSV.csv expense-form.xls

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
['Meals - Lunch', 'Transport - Fuel', 'Entertainment', 'Meals - Dinner', 'Transport - Air', 'Transport - Car Rental', 'Other (travel related)', 'Transport - Taxi', 'Transport - Rail', 'Transport - Parking', 'Miscellaneous (not travel related)', 'Transport - Toll', 'Phone', 'Lodging', 'Meals - Breakfast']
