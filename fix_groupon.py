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



