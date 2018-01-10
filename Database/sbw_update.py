import sqlite3
import openpyxl

print('Starting...')

conn = sqlite3.connect('star_interactive.db')
cur = conn.cursor()

inventory_file = "SBW_InventoryData.xlsx" 
wb = openpyxl.load_workbook(inventory_file)
sheet = wb.active

added_skus = set()
added_count = 0
updated_skus = set()
no_upc = list()

for row in range(2, sheet.max_row + 1):

	item_num = sheet['A' + str(row)].value
	item_sku = sheet['B' + str(row)].value
	item_desc = sheet['C' + str(row)].value
	item_inv = int(sheet['E' + str(row)].value)
	item_upc = sheet['I' + str(row)].value
	item_info = (item_num, item_sku, item_desc, item_inv, item_upc)
	
	if item_upc == '':
		item_upc = str(0)
		no_upc.append(item_sku)

	try:
		cur.execute("INSERT INTO sbw_inventory VALUES (?, ?, ?, ?, ?)", item_info)
		#conn.commit()
		print('Added', item_sku)
		added_skus.add(item_sku)
		print('Added {} to database.'.format(item_sku))
		added_skus.add(item_sku)
		added_count += 1

	except sqlite3.IntegrityError:
		cur.execute("SELECT item_inv FROM sbw_inventory WHERE item_sku = ?", (item_sku,))
		prev_inv = int(cur.fetchone()[0])
		cur.execute("UPDATE sbw_inventory SET item_inv = ?  WHERE item_sku = ?", (item_inv, item_sku))
		print("{} {} ---> {}".format(item_sku, prev_inv, item_inv))

		if prev_inv != item_inv:
			updated_skus.add(item_sku)

print('Updated.')
print('SKUs added:', added_count)
for sku in added_skus:
	print(sku)
print("---")
print('SKUs updated:', len(updated_skus))
for sku in updated_skus:
	print(sku)

conn.commit()
cur.close()
conn.close()

		

