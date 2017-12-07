# SI_OrderOpener

A GUI application to automate the opening of orders to Lean Supply warehouse. Takes order sheets from CommerceHub (Walmart, Best Buy, Staples) and Groupon as inputs and combines them into a single output sheet formatted in a way that Lean Supply will accept. Connects to a MySQL database to fetch information on SKU and UPC codes.
 
## To-Do:
- Clean up GUI interface (add option to reset input sheets)
- Fix bugs with CommerceHub inputs
- Create function to detect input format as .csv and convert to .xlsx
- Add menubar options to update database
- Create MySQL table for Groupon descriptions and associated SKUs, add option to update to menubar

