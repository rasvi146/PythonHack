import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

# Set the URL for the result list
url = "https://results.vasaloppet.se/2024/"

# Send a GET request to the URL
response = requests.get(url)

# Check if the request was successful
if response.status_code == 200:
    # Parse the HTML content
    soup = BeautifulSoup(response.text, 'html.parser')



"""
    # Create a new Excel workbook and select the active sheet
    workbook = Workbook()
    sheet = workbook.active

    # Find and loop through the result list items
    result_items = soup.find_all('your_element_tag_here', class_='your_class_name_here')

    for result_item in result_items:
        # Extract data from the result item and add it to the sheet
        # Modify the code according to the actual HTML structure
        # Example:
        name = result_item.find('span', class_='name').text
        time = result_item.find('span', class_='time').text

        # Append the data to the sheet
        sheet.append([name, time])

    # Save the workbook
    workbook.save('Vasaloppet_Results.xlsx')
"""


else:
    print(f"Failed to retrieve the page. Status code: {response.status_code}")
