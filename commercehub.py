import openpyxl
import csv
import time
from dicts import *
from fix_groupon import *
import datetime

# initial state for logic gates
COMMERCEHUB = False
GROUPON = True
STAPLES = False

# inputs - drop enter in the spreadsheets (commercehub, groupon, staples)
#input_file = input("Enter the CommerceHub file to be formatted:")
commerce_file = "CSV 11-03-2017.xlsx"
groupon_file = 'Groupon-11-07-17.xlsx'
staples_file = 'Staples test 11-07-2017.xlsx'

#input_wb = openpyxl.load_workbook(input_file)
input_wb = openpyxl.load_workbook(groupon_file)
#input_wb = openpyxl.load_workbook(staples_file)
input_sheet = input_wb.active

output_wb = openpyxl.Workbook()
output_sheet = output_wb.active

col_width = 20

# write columns for new output wb, set col. width
format_file = "CSV 11-03-2017.xlsx" # Always use this spreadsheet as the template for the column names/formatting
format_wb = openpyxl.load_workbook(format_file)
format_sheet = format_wb.active
for col in range(1, format_sheet.max_column + 1):
	col_letter = openpyxl.cell.cell.get_column_letter(col)
	output_sheet[col_letter + str(1)] = format_sheet[col_letter + str(1)].value
	output_sheet.column_dimensions[col_letter].width = col_width

if COMMERCEHUB == True:
	print('Adding CommerceHub')
	# write vals
	for row in range(2, input_sheet.max_row + 1):
		# get information that will be reused multiple times
		for col in range(1, format_sheet.max_column + 1):
			col_letter = openpyxl.cell.cell.get_column_letter(col)
			# RULES
			if col_letter == 'C':
				if input_sheet[col_letter + str(row)].value == 'N/A':
					output_sheet[col_letter + str(row)] = None
			elif commercehub_cols[col_letter] == None or commercehub_cols[col_letter] == NA or commercehub_cols[col_letter] == 'US':
				  output_sheet[col_letter + str(row)] = commercehub_cols[col_letter]
			else:
				output_sheet[col_letter + str(row)] = input_sheet[commercehub_cols[col_letter] + str(row)].value
		order_dates(row, output_sheet)

elif GROUPON == True:
	print('Adding Groupon')
	# write vals
	for row in range(2, input_sheet.max_row + 1):
		# get information that will be reused multiple times
		for col in range(1, format_sheet.max_column):
			try:
				col_letter = openpyxl.cell.cell.get_column_letter(col)
				# RULES
				if (groupon_cols[col_letter] == None or groupon_cols[col_letter] == NA or groupon_cols[col_letter] == 'US'
				 or groupon_cols[col_letter] == 'NEED DATE' or groupon_cols[col_letter] == 'Groupon'
				 or groupon_cols[col_letter] == 'Canada Post - Expedited Parcel' or groupon_cols[col_letter] == 'CA'
				 or groupon_cols[col_letter] == 'IGNORE ME'):
					  #print(col_letter + str(row))
					  output_sheet[col_letter + str(row)] = groupon_cols[col_letter]
				else:
					print(col_letter + str(row))
					output_sheet[col_letter + str(row)] = input_sheet[groupon_cols[col_letter] + str(row)].value
			except Exception:
				pass
		grab_skus_upc(row, output_sheet)
		order_dates(row, output_sheet)

elif STAPLES == True:
	print('Adding Staples')
	# write vals
	for row in range(2, input_sheet.max_row + 1):
		# get information that will be reused multiple times
		for col in range(1, format_sheet.max_column):
			try:
				col_letter = openpyxl.cell.cell.get_column_letter(col)
				# RULES
				if (staples_cols[col_letter] == None or staples_cols[col_letter] == NA or staples_cols[col_letter] == 'US'
				 or staples_cols[col_letter] == 'NEED DATE' or staples_cols[col_letter] == 'Staples'
				 or staples_cols[col_letter] == 'Canada Post - Expedited Parcel' or staples_cols[col_letter] == 'CA'
				 or staples_cols[col_letter] == 'IGNORE ME'):
					  #print(col_letter + str(row))
					  output_sheet[col_letter + str(row)] = staples_cols[col_letter]
				else:
					print(col_letter + str(row))
					output_sheet[col_letter + str(row)] = input_sheet[staples_cols[col_letter] + str(row)].value
			except Exception:
				pass
		#grab_skus_upc(row, output_sheet) <------ Aroma Oils Fix
		order_dates(row, output_sheet)

output_wb.save("FORMATTED_" + staples_file)
print('Done.')
