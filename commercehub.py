import openpyxl
import csv
import time
from dicts import *
from fix_groupon import *
import datetime
import sys
import pymysql

# initial state for logic gates
COMMERCEHUB = False
GROUPON = True
STAPLES = True

# connect to db for SKUs and UPCs
conn = pymysql.connect(host='127.0.0.1', user='root', passwd='Greengiant90',
						db='mysql', charset='utf8')
cur = conn.cursor()
cur.execute("USE star_interactive")

# inputs - drop enter in the spreadsheets (commercehub, groupon, staples)
#input_file = input("Enter the CommerceHub file to be formatted:")
commerce_file = "CSV 11-03-2017.xlsx"
groupon_file = 'Groupon 11-10-2017.xlsx'
staples_file = 'Staples test 11-07-2017.xlsx'

# check for inputs

if (COMMERCEHUB == False and GROUPON == False and STAPLES == False):
	print('No input files entered.')
	sys.exit()

output_wb = openpyxl.Workbook()
output_sheet = output_wb.active

# write columns for new output wb, set col. width
col_width = 20
format_file = "CSV 11-03-2017.xlsx" # Always use this spreadsheet as the template for the column names/formatting
format_wb = openpyxl.load_workbook(format_file)
format_sheet = format_wb.active
for col in range(1, format_sheet.max_column + 1):
	col_letter = openpyxl.cell.cell.get_column_letter(col)
	output_sheet[col_letter + str(1)] = format_sheet[col_letter + str(1)].value
	output_sheet.column_dimensions[col_letter].width = col_width

final_col = format_sheet.max_column + 1

# row offset for multiple input sheets
offset = 0
# error_rows for final checking
error_rows = set()

if COMMERCEHUB == True:
	print('Adding CommerceHub')
	input_wb = openpyxl.load_workbook(commerce_file)
	input_sheet = input_wb.active
	last_row = input_sheet.max_row
	print('SKUs:', last_row - 1)
	# write vals
	for row in range(2, last_row + 1):
		# get information that will be reused multiple times
		for col in range(1, final_col):
			col_letter = openpyxl.cell.cell.get_column_letter(col)
			# RULES
			if col_letter == 'C':
				if input_sheet[col_letter + str(row)].value == 'N/A':
					output_sheet[col_letter + str(row)] = None
			elif commercehub_cols[col_letter] == None or commercehub_cols[col_letter] == NA or commercehub_cols[col_letter] == 'US':
				  output_sheet[col_letter + str(row)] = commercehub_cols[col_letter]
			else:
				output_sheet[col_letter + str(row)] = str(input_sheet[commercehub_cols[col_letter] + str(row)].value)
		order_dates(row, output_sheet)
		mysql_lookup(row, output_sheet, cur)
		check_errors(row, final_col, output_sheet, error_rows)
	offset += last_row - 1

if GROUPON == True:
	print('Adding Groupon')
	input_wb = openpyxl.load_workbook(groupon_file)
	input_sheet = input_wb.active
	last_row = input_sheet.max_row
	print('SKUs:', last_row - 1)
	# write vals
	for row in range(2, last_row + 1):
		# get information that will be reused multiple times
		for col in range(1, final_col):
			try:
				col_letter = openpyxl.cell.cell.get_column_letter(col)
				# RULES
				if (groupon_cols[col_letter] == None or groupon_cols[col_letter] == NA or groupon_cols[col_letter] == 'US'
				 or groupon_cols[col_letter] == 'NEED DATE' or groupon_cols[col_letter] == 'Groupon'
				 or groupon_cols[col_letter] == 'Canada Post - Expedited Parcel' or groupon_cols[col_letter] == 'CA'
				 or groupon_cols[col_letter] == 'IGNORE ME'):
					  #print(col_letter + str(row))
					  output_sheet[col_letter + str(row + offset)] = groupon_cols[col_letter]
				else:
					#print(col_letter + str(row))
					output_sheet[col_letter + str(row + offset)] = input_sheet[groupon_cols[col_letter] + str(row)].value
			except Exception:
				pass
		grab_skus_upc(row + offset, output_sheet)
		order_dates(row + offset, output_sheet)
		mysql_lookup(row + offset, output_sheet, cur)
		check_errors((row + offset), final_col, output_sheet, error_rows)
	offset += last_row - 1

if STAPLES == True:
	print('Adding Staples')
	input_wb = openpyxl.load_workbook(staples_file)
	input_sheet = input_wb.active
	last_row = input_sheet.max_row
	print('SKUs:', last_row - 1)
	# write vals
	for row in range(2, last_row + 1):
		# get information that will be reused multiple times
		for col in range(1, final_col):
			try:
				col_letter = openpyxl.cell.cell.get_column_letter(col)
				# RULES
				if (staples_cols[col_letter] == None or staples_cols[col_letter] == NA or staples_cols[col_letter] == 'US'
				 or staples_cols[col_letter] == 'NEED DATE' or staples_cols[col_letter] == 'Staples'
				 or staples_cols[col_letter] == 'Canada Post - Expedited Parcel' or staples_cols[col_letter] == 'CA'
				 or staples_cols[col_letter] == 'IGNORE ME'):
					  #print(col_letter + str(row))
					  output_sheet[col_letter + str(row + offset)] = staples_cols[col_letter]
				else:
					#print(col_letter + str(row))
					output_sheet[col_letter + str(row + offset)] = input_sheet[staples_cols[col_letter] + str(row)].value
			except Exception:
				pass
		#grab_skus_upc(row, output_sheet) <------ Aroma Oils Fix
		order_dates(row + offset, output_sheet)
		mysql_lookup(row + offset, output_sheet, cur)
		check_errors((row + offset), final_col, output_sheet, error_rows)
	offset += last_row - 1

today = datetime.date.today()
output_file = today.strftime("%m-%d-%Y") + " ORDERS.xlsx"
output_wb.save(output_file)
print('Done.')
print('Total SKUs:', output_sheet.max_row - 1)
print('-----')
for row in error_rows:
	print('WARNING! Potential error on row', row)

cur.close()
conn.close()
