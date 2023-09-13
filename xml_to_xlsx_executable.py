
import pandas as pd
from bs4 import BeautifulSoup
import time

try:
    # Prompt the user to enter the XML file location
    xml_file_path = input("Please Enter the path to the MSM.xml file: ")

    # Prompt the user to enter the XLS file location
    xls_file_path = input("Please Enter the path where you want to store XLS file: ")

    # Parse the XML file
    with open(xml_file_path) as file:
        contents = file.read()
        soup = BeautifulSoup(contents, 'xml')

    # Create empty lists to store the extracted values
    codes = []
    sub_code_modifiers = []
    amounts = []
    counts = []

    # Find all MSMDetail elements
    msm_details = soup.find_all('MSMDetail')

    # Loop through each MSMDetail element
    for detail in msm_details:
        # Find the MiscellaneousSummaryCode element and extract its value
        code = detail.find('MiscellaneousSummaryCode').text
        codes.append(code)

        # Find the MiscellaneousSummarySubCodeModifier element and extract its value
        sub_code_element = detail.find('MiscellaneousSummarySubCodeModifier')
        sub_code = sub_code_element.text

        # Map numeric sub_code values to their corresponding names
        if sub_code == "600":
            sub_code = "Debit"
        elif sub_code == "1800":
            sub_code = "Wex"
        elif sub_code == "800":
            sub_code = "Discover"
        elif sub_code == "1451":
            sub_code = "Fuelman"
        elif sub_code == "951":
            sub_code = "EBT Food"
        elif sub_code == "1700":
            sub_code = "Voyager"
        elif sub_code == "952":
            sub_code = "EBT Cash Benefit"
        elif sub_code == "2600":
            sub_code = "Fleet One"
        elif sub_code == "1100":
            sub_code = "Mastercard"
        elif sub_code == "1101":
            sub_code = "Mastercard Fleet"
        elif sub_code == "1518":
            sub_code = "CK Gift Card"
        elif sub_code == "1600":
            sub_code = "Visa"
        elif sub_code == "1602":
            sub_code = "Visa Fleet"
        elif sub_code == "1113":
            sub_code = "Easy Pay"
        elif sub_code == "100":
            sub_code = "AMEX"

        sub_code_modifiers.append(sub_code)

        # Find the MiscellaneousSummaryAmount element and extract its value as a float
        amount_text = detail.find('MiscellaneousSummaryAmount').text
        amount = float(amount_text) if amount_text else 0.0  # Convert to float if not empty, else set to 0.0
        amounts.append(amount)

        # # Find the MiscellaneousSummaryAmount element and extract its value
        # count_text = detail.find('MiscellaneousSummaryCount').text
        # count = float(count_text) if count_text else 0.0  # Convert to float if not empty, else set to 0.0
        # counts.append(count)

        # Find the MiscellaneousSummaryCount element and extract its value if it exists
        count_element = detail.find('MiscellaneousSummaryCount')
        count = count_element.text if count_element is not None else ''
        counts.append(count)



    # Create a pandas DataFrame with the extracted values
    data = {
        'MiscellaneousSummaryCode': codes,
        'MiscellaneousSummarySubCodeModifier': sub_code_modifiers,
        'MiscellaneousSummaryAmount': amounts,
        'MiscellaneousSummaryCount': counts
    }
    df = pd.DataFrame(data)

    # Perform the cleaning: Keep only rows where the SubCodeModifier is in the desired list
    desired_values = ['Debit', 'Wex', 'Discover', 'Fuelman', 'EBT Food', 'Voyager', 'EBT Cash Benefit', 'Fleet One', 'Mastercard', 'Mastercard Fleet', 'CK Gift Card', 'Visa', 'Visa Fleet', 'Easy Pay', 'AMEX']
    cleaned_df = df[df['MiscellaneousSummarySubCodeModifier'].isin(desired_values)]


     # Create a pivot table
    pivot_table = cleaned_df.pivot_table(index='MiscellaneousSummaryCode', columns='MiscellaneousSummarySubCodeModifier', values=['MiscellaneousSummaryAmount', 'MiscellaneousSummaryCount'], aggfunc={'MiscellaneousSummaryAmount':'sum', 'MiscellaneousSummaryCount':'sum'}, fill_value=0)

    # Save the pivot table to an Excel file
    pivot_table.to_excel(xls_file_path)

    # Display a completion message
    print("Your spreadsheet is ready - check it in the folder.")

    # Add a 30-second delay
    time.sleep(10)

except Exception as e:
    # Handle exceptions and display an error message
    print(f"An error occurred: {str(e)}")
    time.sleep(10)

# Close the script
