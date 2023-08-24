# extracts refunds but for some reason in the end of the table

import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import NamedStyle

# Specify the path to your XML file
xml_file_path = "/Users/anna/Desktop/CK_Rep/Test/NAXML-08-23.xml"

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
    "Transaction ID", "MOP", "Description", "Gallons Used",
    "Fuel $ ", "Merch Sales $", "Tax", "Credit", "Debit", "Cash", "EBT", "Discount Amount", "Discount Reason"
]
sheet.append(headers)

# Define the number format for the "Fuel $" column
fuel_format = NamedStyle(name="custom_format", number_format="0,000")

# Iterate over each SaleEvent element
for sale_event in sale_events:
    # Find the elements within each SaleEvent
    transaction_id = sale_event.find(".//ns0:TransactionID", namespace)
    card_descr = sale_event.find(".//radiant:ShorterDescription", namespace)
    description = sale_event.find(".//ns0:Description", namespace)
    sales_quantity = sale_event.find(".//ns0:SalesQuantity[@uom='gallonUS']", namespace)
    fuel_amounts = sale_event.findall(".//ns0:SalesAmount", namespace)  # Find all SalesAmount elements
    tender_sub_code = sale_event.find(".//ns0:TenderSubCode", namespace)
    tax_collected = sale_event.find(".//ns0:TransactionTotalTaxSalesAmount", namespace)
    tender_amount = sale_event.find(".//ns0:TenderAmount", namespace)
    discount_amount = sale_event.find(".//ns0:DiscountAmount", namespace)
    discount_reason = sale_event.find(".//ns0:DiscountReason", namespace)

    # Extract the values for each row
    row_values = [
        transaction_id.text if transaction_id is not None else "",
        card_descr.text if card_descr is not None else "Cash",
        description.text if description is not None else "",
        sales_quantity.text if sales_quantity is not None else "",
        fuel_amounts[0].text if fuel_amounts else "",
        fuel_amounts[1].text if len(fuel_amounts) > 1 else "",
        tax_collected.text if tax_collected is not None and tax_collected.text != "0" else "",  # Set default value to empty string if "0"
        "",  # Placeholder for Credit amount (initialize to empty string)
        "",  # Placeholder for Debit amount (initialize to empty string)
        "",  # Placeholder for Cash amount (initialize to empty string)
        "",  # Placeholder for EBT amount (initialize to empty string)
        discount_amount.text if discount_amount is not None else "",
        discount_reason.text if discount_reason is not None else "",
    ]

    # Find the payment methods and their corresponding amounts
    payment_methods = sale_event.findall(".//ns0:TenderSubCode", namespace)
    for method in payment_methods:
        method_value = method.text.strip().lower()
        if method_value == "credit":
            row_values[headers.index("Credit")] = tender_amount.text if tender_amount is not None else ""
        elif method_value == "debit":
            row_values[headers.index("Debit")] = tender_amount.text if tender_amount is not None else ""
        elif method_value == "cash":
            row_values[headers.index("Cash")] = tender_amount.text if tender_amount is not None else ""
        elif method_value == "ebt":
            row_values[headers.index("EBT")] = tender_amount.text if tender_amount is not None else ""

    # Write the row to the Excel sheet
    sheet.append(row_values)

# Apply the number format to the "Fuel $" column
fuel_column_number = headers.index("Fuel $ ")
for cell in sheet[chr(65 + fuel_column_number)]:
    cell.style = fuel_format

# Find all RefundEvent elements
refund_events = root.findall(".//ns0:RefundEvent", namespace)

# Iterate over each RefundEvent element
for refund_event in refund_events:
    # Find the elements within each RefundEvent
    transaction_id = refund_event.find(".//ns0:TransactionID", namespace)
    tender_amount = refund_event.find(".//ns0:TenderAmount", namespace)

    # Extract the values for each row
    row_values = [
        transaction_id.text if transaction_id is not None else "",
        "",  # Placeholder for MOP (initialize to empty string)
        "",  # Placeholder for Description (initialize to empty string)
        "",  # Placeholder for Gallons Used (initialize to empty string)
        "",  # Placeholder for Fuel $ (initialize to empty string)
        "",  # Placeholder for Merch Sales $ (initialize to empty string)
        "",  # Placeholder for Tax (initialize to empty string)
        "",  # Placeholder for Credit amount (initialize to empty string)
        "",  # Placeholder for Debit amount (initialize to empty string)
        "",  # Placeholder for Cash amount (initialize to empty string)
        "",  # Placeholder for EBT amount (initialize to empty string)
    ]

    # Update the tender amount for the corresponding MOP column
    row_values[headers.index("MOP")] = "Refund"
    row_values[headers.index("Credit")] = tender_amount.text if tender_amount is not None else ""

    # Write the row to the Excel sheet
    sheet.append(row_values)


# Save the workbook to a file
excel_file_path = "/Users/anna/Desktop/CK_Rep/Test/NAxml_test_8.23_REFUNDS_3.xlsx"
workbook.save(excel_file_path)

print("Data has been written to the Excel file.")
