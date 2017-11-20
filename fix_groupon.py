def grab_skus_upc(row, output_sheet):
	product_desc = output_sheet['AM'+ str(row)].value
	#print(product_desc)
	if product_desc.endswith('(14 Pack)'):
		# UPC
		output_sheet['CQ'+ str(row)] = '743724486834'
		# SKU
		output_sheet['CR'+ str(row)] = 'AO-14'

	elif product_desc.endswith('(6 Pack)'):
		# UPC
		output_sheet['CQ'+ str(row)] = '743724486827'
		# SKU
		output_sheet['CR'+ str(row)] = 'AO-6'

	elif product_desc.endswith('Set'):
		# UPC
		output_sheet['CQ'+ str(row)] = '743724487237'
		# SKU
		output_sheet['CR'+ str(row)] = 'AO-8'

import datetime
import openpyxl
import pymysql

def order_dates(row, output_sheet):
	today = datetime.date.today()
	tomorrow = datetime.date.today() + datetime.timedelta(days=1)
	output_sheet['BH'+ str(row)] = today.strftime("%m-%d-%Y")
	output_sheet['BN'+ str(row)] = tomorrow.strftime("%m-%d-%Y")

def check_errors(row, final_col, output_sheet, error_rows):
	for col in range(1, final_col):
		col_letter = openpyxl.cell.cell.get_column_letter(col)
		if (output_sheet[col_letter + str(row)].value == 'IGNORE ME' or output_sheet[col_letter + str(row)].value == 'N/A'
			or output_sheet[col_letter + str(row)].value == '0'):
			error_rows.add(row)

def mysql_lookup(row, output_sheet, cur):
	row_sku = output_sheet["CR" + str(row)].value
	query = """SELECT upc FROM lean_supply WHERE sku = "{}"; """
	try:
		cur.execute(query.format(row_sku))
		row_upc = cur.fetchone()[0]
		#print(query.format(row_sku))
		print('Grabbed UPC from database for', row_sku)
		output_sheet["CQ" + str(row)] = str(row_upc)
	except Exception as e:
		print(e)









