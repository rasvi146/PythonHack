from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
import time

###### To enable saving to xml ######
from bs4 import BeautifulSoup
#####################################

###### for debugging ######
import sys
###########################

####### Initiation #######
# Allow å,ä,ö
sys.stdout.reconfigure(encoding='utf-8')

# Start a new Selenium WebDriver session
driver = webdriver.Chrome()

# Define wait, for the page to load and elements to be present
wait = WebDriverWait(driver, 10)

# Specify the year and event you want to select
selected_year = '2024'
selected_event_value = 'Öppet Spår söndag'
###########################


def initiate_advanced_search():

    driver.get("https://results.vasaloppet.se/2024/")

    # Locate the "Advanced Search" section and un-collapse it
    advanced_search_section = driver.find_element(By.ID, "collapse-advanced")
    driver.execute_script("arguments[0].setAttribute('class', 'panel-collapse collapse in')", advanced_search_section)

    # Find the year dropdown in the "Advanced Search" section
    year_dropdown = Select(driver.find_element(By.ID, "advanced-search-event_main_group"))

    # Choose the selected year
    year_dropdown.select_by_value(selected_year)

    # Wait for the page to load
    time.sleep(1)  # You can adjust the waiting time as needed

    # Find the event dropdown in the "Advanced Search" section
    event_dropdown = driver.find_element(By.ID, "advanced-search-event")

    # Create a Select object for the dropdown
    select_event = Select(event_dropdown)

    # Choose the event "Öppet Spår söndag" by its visible text
    select_event.select_by_visible_text(selected_event_value)

    # Wait for 1 second to allow page to load
    time.sleep(1)

    # Find the "Results/Page" dropdown in the "Advanced Search" section
    results_per_page_dropdown = driver.find_element(By.ID, "advanced-num_results")

    # Create a Select object for the dropdown
    select_results_per_page = Select(results_per_page_dropdown)

    # Choose the option "100" by its visible text
    select_results_per_page.select_by_visible_text("100")

    # Submit the form
    submit_button = wait.until(EC.element_to_be_clickable((By.ID, 'advanced-submit')))
    submit_button.click()

    time.sleep(1)


def collect_participants_page():

    # Wait for the new page to load
    wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'list-active')))

    # Get the page source after the form submission
    page_source = driver.page_source

    # Use BeautifulSoup to parse the HTML
    soup = BeautifulSoup(page_source, 'html.parser')

    # Create a matrix to store participant information
    participant_matrix = []

    # Find all participant elements on the page
    participant_elements = soup.find_all('li', class_='list-group-item row') + \
                           soup.find_all('li', class_='list-active list-group-item row') + \
                           soup.find_all('li', class_='list-active event-ÖSS_HCH8NDMR2400 list-group-item row') + \
                           soup.find_all('li', class_='event-ÖSS_HCH8NDMR2400 list-group-item row')

    # Iterate through each participant element and store the extracted information in the matrix
    for participant_element in participant_elements:
        # Extract name
        name_element = participant_element.find('h4', class_='list-field type-fullname')
        name = name_element.text.strip()[:-6] if name_element else None

        # Extract number
        number_element = participant_element.find('div', class_='list-field type-field')
        number = number_element.text.replace('Number', '').strip() if number_element else None

        # Extract finish time
        finish_time_element = participant_element.find('div', class_='right list-field type-time')
        finish_time = finish_time_element.text.replace('Finish', '').strip() if finish_time_element else None

        # Append the participant information to the matrix
        participant_matrix.append([name, number, finish_time])

    # # Print the matrix
    # for participant_info in participant_matrix:
    #         print(participant_info)
    # print(participant_matrix[1])
    # print(participant_matrix[-1])
    # print('')

    return participant_matrix


def save_to_excel(participant_matrix_total):
    # Create a new Excel workbook and add a worksheet
    wb = Workbook()
    ws = wb.active

    # Write header row
    ws.append(['Name', 'Number', 'Finish Time'])

    # Write participant data
    for participant_matrix in participant_matrix_total:
        for participant_info in participant_matrix:
            ws.append(participant_info)

    # Save the workbook
    wb.save('vasaloppet_results_' + str(selected_year) + '.xlsx')


def main():
    initiate_advanced_search()

    # participant_matrix

    page_number = 1  # Initialize page number

    participant_matrix_total = []

    while True:
        participant_matrix = collect_participants_page()

        participant_matrix_total.append(participant_matrix)

        # Check if there is a next page button
        button_name = 'li.pages-nav-button a[href*="page={}"]'.format(page_number + 1)
        next_page_button = driver.find_elements(By.CSS_SELECTOR, button_name)

        if next_page_button:
            page_number += 1

            print(f'page number: {page_number}')
            # Click the next page button
            next_page_button[0].click()  # Click the first button found
            time.sleep(1)  # Adjust the waiting time as needed
        else:
            break

    # Save to Excel
    save_to_excel(participant_matrix_total)

    # Close the browser window
    driver.quit()


if __name__ == '__main__':
    main()
