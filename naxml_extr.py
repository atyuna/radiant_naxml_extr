import xml.etree.ElementTree as ET
from openpyxl import Workbook

# Specify the path to your XML file
xml_file_path = "/Users/anna/Desktop/CK_Rep/NAXML-POS.xml"

# Define the namespace dictionary with the prefix and URI
namespace = {
    "ns0": "http://www.naxml.org/POSBO/Vocabulary/2003-10-16",
    "radiant": "http://www.radiantsystems.com/NAXML-Extension",
    "xsi": "http:www.w3.org/20011/XMLSchema-instance",
    "xmime": "http://www.w3.org/2005/05/xmlmime"
}

# Parse the XML file
tree = ET.parse(xml_file_path)
root = tree.getroot()

# Find all SaleEvent elements
sale_events = root.findall(".//ns0:SaleEvent", namespace)

# Create a new Excel workbook and get the active sheet
workbook = Workbook()
sheet = workbook.active

# Write the column headers
headers = [
    "Transaction ID", "Fuel Grade", "Gallons Used",
    "Fuel $ ", "Merch Sales $", "Method of Payment", "Tax", "Total"
]
sheet.append(headers)

# Iterate over each SaleEvent element
for sale_event in sale_events:
    # Find the elements within each SaleEvent
    transaction_id = sale_event.find(".//ns0:TransactionID", namespace)
    description = sale_event.find(".//ns0:Description", namespace)
    sales_quantity = sale_event.find(".//ns0:SalesQuantity[@uom='gallonUS']", namespace)
    fuel_amounts = sale_event.findall(".//ns0:SalesAmount", namespace)  # Find all SalesAmount elements
    tender_sub_code = sale_event.find(".//ns0:TenderSubCode", namespace)
    tax_collected = sale_event.find(".//ns0:TransactionTotalTaxSalesAmount", namespace)
    tender_amount = sale_event.find(".//ns0:TenderAmount", namespace)

    # Extract the values for each row
    row_values = [
        transaction_id.text if transaction_id is not None else "-",
        # fuel_grade_id.text if fuel_grade_id is not None else "Fuel Grade ID not found",
        description.text if description is not None else "-",
        sales_quantity.text if sales_quantity is not None else "0",
        fuel_amounts[0].text if fuel_amounts else "0",
        fuel_amounts[1].text if len(fuel_amounts) > 1 else "0",
        tender_sub_code.text if tender_sub_code is not None else "0",
        tax_collected.text if tax_collected is not None else "0",
        tender_amount.text if tender_amount is not None else "0"
    ]

    # Write the row to the Excel sheet
    sheet.append(row_values)

# Save the workbook to a file
excel_file_path = "/Users/anna/Desktop/CK_Rep/NAxml_Reconsiliation.xlsx"
workbook.save(excel_file_path)

print("Data has been written to the Excel file.")
