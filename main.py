import pandas as pd
import requests
from datetime import datetime

# Define the file path and sheet name
excel_file_path = 'Book1.xlsx'  # Ensure this path is correct
sheet_name = 'Sheet1'  # Ensure this matches the sheet name in your Excel file

# Load the Excel file
df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

# Check if the required column exists
if 'walletAddress' not in df.columns:
    raise ValueError("The required column 'walletAddress' does not exist in the Excel sheet.")

# Define the base URL for the GET request
base_url = 'https://kx58j6x5me.execute-api.us-east-1.amazonaws.com/scroll/wallet-points'

# List to store the response data
response_data_list = []

# Iterate over each row in the DataFrame and make GET requests
for index, row in df.iterrows():
    # Extract the parameter from the current row
    wallet_address = row['walletAddress']

    # Construct the complete URL
    request_url = f"{base_url}?walletAddress={wallet_address}"

    # Make the GET request
    response = requests.get(request_url)

    # Check if the request was successful
    if response.status_code == 200:
        try:
            # Parse the JSON response
            json_response = response.json()

            # Convert the list of dictionaries to a single dictionary
            if isinstance(json_response, list):
                data_dict = {'points': item['points'] for item in json_response}
                # Or, to simply get the 'points' value from the first item, you can directly assign:
                # data = json_response[0].get('points', 'No data') if json_response else 'No data'
                data = data_dict.get(wallet_address, data_dict['points'])  # Adjust according to your needs
            else:
                data = json_response.get('points')

        except ValueError:
            data = 'Invalid JSON'
    else:
        data = 'Request failed'

    response_data_list.append(round(data))

today = datetime.now()

# Add the response data to a new column in the DataFrame
df[today] = response_data_list

# Save the updated DataFrame back to the Excel file
df.to_excel(excel_file_path, sheet_name=sheet_name, index=False)

print("Data updated successfully in the Excel file.")
