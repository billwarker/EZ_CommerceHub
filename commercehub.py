import openpyxl
import csv
import time
from dicts import *
from fix_groupon import *

COMMERCEHUB = False
GROUPON = True

#input_file = input("Enter the CommerceHub file to be formatted:")
input_file = "CSV 11-03-2017.xlsx"
groupon_file = 'groupon orders2.xlsx'

#input_wb = openpyxl.load_workbook(input_file)
input_wb = openpyxl.load_workbook(groupon_file)
input_sheet = input_wb.active

output_wb = openpyxl.Workbook()
output_sheet =  output_wb.active

col_width = 20

# write columns for new output wb, set col. width
format_wb = openpyxl.load_workbook(input_file)
format_sheet = format_wb.active
for col in range(1, format_sheet.max_column + 1):
	col_letter = openpyxl.cell.cell.get_column_letter(col)
	output_sheet[col_letter + str(1)] = format_sheet[col_letter + str(1)].value
	output_sheet.column_dimensions[col_letter].width = col_width

if COMMERCEHUB == True:
# write vals
	for row in range(2, input_sheet.max_row + 1):
		# get information that will be reused multiple times
		for col in range(1, input_sheet.max_column + 1):
			col_letter = openpyxl.cell.cell.get_column_letter(col)
			# RULES
			if col_letter == 'C':
				if input_sheet[col_letter + str(row)].value == 'N/A':
					output_sheet[col_letter + str(row)] = None
			elif commercehub_cols[col_letter] == None or commercehub_cols[col_letter] == NA or commercehub_cols[col_letter] == 'US':
				  output_sheet[col_letter + str(row)] = commercehub_cols[col_letter]
			else:
				output_sheet[col_letter + str(row)] = input_sheet[commercehub_cols[col_letter] + str(row)].value

	output_wb.save("FORMATTED_" + input_file)
	print('Done.')

elif GROUPON == True:
# write vals
	for row in range(2, input_sheet.max_row + 1):
		# get information that will be reused multiple times
		for col in range(1, input_sheet.max_row):
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

output_wb.save("FORMATTED_" + groupon_file)
print('Done.')
