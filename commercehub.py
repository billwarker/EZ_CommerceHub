import openpyxl
import csv
import time

#input_file = input("Enter the CommerceHub file to be formatted:")
input_file = "CSV 11-03-2017.xlsx"

input_wb = openpyxl.load_workbook(input_file)
input_sheet = input_wb.active

output_wb = openpyxl.Workbook()
output_sheet =  output_wb.active

#null_cols = ['A','D','M','N','O','Q','S','T','U','W','X','AF','AG','AK','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ',
#			'BA','BB','BC','BE','BF','BG','BL','BO','BP','BS','BT','BU','BV','CG','CI','CJ','CK','CL','CM','CN','CO','CP']
#NA_cols = ['I', 'J', 'K', 'AC', 'AD', 'CE']
#customer_name_cols = ['L', 'F', 'Z', 'AE', 'CB', 'CC', 'CD']
#address_cols = ['V','BR']
#city_cols = ['E','Y','AB','BW']
#postal_code_cols = ['P', 'AJ', 'CF']
#province_cols = ['R', 'AL', 'BX', 'CH']
#country_cols = ['G', 'H', 'AA']
#customer_order_number_cols = ['AI', 'BI', 'BK']
#customer_number_col = 'BY'

col_width = 20

# write columns for new output wb, set col. width
for col in range(1, input_sheet.max_column + 1):
	col_letter = openpyxl.cell.cell.get_column_letter(col)
	output_sheet[col_letter + str(1)] = input_sheet[col_letter + str(1)].value
	output_sheet.column_dimensions[col_letter].width = col_width

# write vals
for row in range(2, input_sheet.max_row + 1):
	# get information that will be reused multiple times
	customer_name = input_sheet['L' + str(row)].value
	address = input_sheet['B' + str(row)].value
	city = input_sheet['E' + str(row)].value
	postal_code = input_sheet['P' + str(row)].value
	province = input_sheet['R' + str(row)].value
	country = input_sheet['G' + str(row)].value
	customer_order_number = input_sheet['AI' + str(row)].value
	customer_number = 'US'

	for col in range(1, input_sheet.max_column + 1):
		col_letter = openpyxl.cell.cell.get_column_letter(col)

		# RULES
		if col_letter == 'C':
			if input_sheet[col_letter + str(row)].value == 'N/A':
				output_sheet[col_letter + str(row)] = None
			else:
				output_sheet[col_letter + str(row)] = str(input_sheet[col_letter + str(row)].value)

		# REUSED INFORMATION COLS

		elif (col_letter == 'A' or col_letter == 'D' or col_letter == 'M' or col_letter == 'N' or col_letter == 'O'
			or col_letter == 'Q' or col_letter == 'S' or col_letter == 'T' or col_letter == 'U' or col_letter == 'W'
			or col_letter == 'X' or col_letter == 'AF' or col_letter == 'AG' or col_letter == 'AK' or col_letter == 'AN'
			or col_letter == 'AO' or col_letter == 'AP' or col_letter == 'AQ' or col_letter == 'AR' or col_letter == 'AS'
			or col_letter == 'AT' or col_letter == 'AU' or col_letter == 'AV' or col_letter == 'AW' or col_letter == 'AX'
			or col_letter == 'AY' or col_letter == 'AZ' or col_letter == 'BA' or col_letter == 'BB' or col_letter == 'BC'
			or col_letter == 'BE' or col_letter == 'BF' or col_letter == 'BG' or col_letter == 'BL' or col_letter == 'BO'
			or col_letter == 'BP' or col_letter == 'BS' or col_letter == 'BT' or col_letter == 'BU' or col_letter == 'BV'
			or col_letter == 'CG' or col_letter == 'CI' or col_letter == 'CJ' or col_letter == 'CK' or col_letter == 'CL'
			or col_letter == 'CM' or col_letter == 'CN' or col_letter == 'CO' or col_letter == 'CP' or col_letter == 'BJ'):
			output_sheet[col_letter + str(row)] = None

		elif (col_letter == 'I' or col_letter == 'J' or col_letter == 'K' or col_letter == 'AC' or col_letter == 'AD'
			or col_letter == 'CE') :
			output_sheet[col_letter + str(row)] = 'NA'

		elif (col_letter == 'L' or col_letter == 'F' or col_letter == 'Z' or col_letter == 'AE' or col_letter == 'CB' 
			or col_letter == 'CC' or col_letter == 'CD'):
			output_sheet[col_letter + str(row)] = customer_name

		elif (col_letter == 'V' or col_letter == 'BR'):
			output_sheet[col_letter + str(row)] = address

		elif (col_letter == 'E' or col_letter == 'Y' or col_letter == 'AB' or col_letter == 'BW'):
			output_sheet[col_letter + str(row)] = city

		elif (col_letter == 'P' or col_letter == 'AJ' or col_letter == 'CF'):
			output_sheet[col_letter + str(row)] = postal_code

		elif (col_letter == 'R' or col_letter == 'AL' or col_letter == 'BX' or col_letter == 'CH'):
			output_sheet[col_letter + str(row)] = province

		elif (col_letter == 'G' or col_letter == 'H' or col_letter == 'AA'):
			output_sheet[col_letter + str(row)] = country

		elif (col_letter == 'AI' or col_letter == 'BI' or col_letter == 'BK'):
			output_sheet[col_letter + str(row)] = customer_order_number

		elif col_letter == 'BY':
			output_sheet[col_letter + str(row)] = customer_number

		# EVERYTHING ELSE
		else:
			output_sheet[col_letter + str(row)] = str(input_sheet[col_letter + str(row)].value)

output_wb.save("FORMATTED_" + input_file)
print('Done.')