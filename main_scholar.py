import csv
import openpyxl
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from scholar_2 import PaperScraper
from tor_proxy import TorProxy  # Import the TorProxy class

# Load the Excel file
workbook = openpyxl.load_workbook('CoICT Google Scholar.xlsx')
sheet = workbook['CoICT']  # Replace 'CoICT' with your actual sheet name

registered_hyperlinks = []

# Iterate through each row and check the status column
for row in sheet.iter_rows(min_row=2, min_col=1):
    status_cell = row[5]  # Assuming the status column is at index 5 (column F)
    if status_cell.value == 'Registered' and status_cell.hyperlink:
        hyperlink_address = status_cell.hyperlink.target + '&view_op=list_works&sortby=pubdate'
        name_cell = row[1].value
        registered_hyperlinks.append([name_cell, hyperlink_address])

workbook.close()

# Initialize TorProxy
tor_proxy = TorProxy()
proxy_ip = tor_proxy.get_ip()

# Configure Chrome options for the website
brave_path = "C:\\Program Files\\BraveSoftware\\Brave-Browser\\Application\\brave.exe"
options = Options()
options.binary_location = brave_path
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--start-maximized')
options.add_argument(f'--proxy-server={proxy_ip}')

# Open the Chrome WebDriver using TOR proxy
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

def extract_citation_metrics(driver):
    metrics = {'CITATIONS': 'N/A', 'H_INDEX': 'N/A', 'I10_INDEX': 'N/A'}
    try:
        table = driver.find_element(By.ID, "gsc_rsb_st")
        rows = table.find_elements(By.XPATH, ".//tbody/tr")
        metrics['CITATIONS'] = rows[0].find_elements(By.CLASS_NAME, 'gsc_rsb_std')[0].text
        metrics['H_INDEX'] = rows[1].find_elements(By.CLASS_NAME, 'gsc_rsb_std')[0].text
        metrics['I10_INDEX'] = rows[2].find_elements(By.CLASS_NAME, 'gsc_rsb_std')[0].text
    except Exception as e:
        print(f"Error extracting citation metrics: {e}")
    return metrics

paper_details = []
paper_scraper = PaperScraper()

for hyperlink in registered_hyperlinks:
    print(f"Processing: {hyperlink[0]}")
    driver.get(hyperlink[1])
    citation_metrics = extract_citation_metrics(driver)

    while True:
        try:
            show_more_button = driver.find_element(By.ID, "gsc_bpf_more")
            if show_more_button.is_enabled():
                driver.execute_script("arguments[0].scrollIntoView(true);", show_more_button)
                show_more_button.click()
                time.sleep(2)
            else:
                break
        except:
            break

    elements = driver.find_elements(By.XPATH, '//a[@class="gsc_a_at"]')
    span_elements = driver.find_elements(By.XPATH, '//span[@class="gsc_a_h gsc_a_hc gs_ibl"]')
    cite_elements = driver.find_elements(By.XPATH, '//a[@class="gsc_a_ac gs_ibl"]')
    
    
    # Start the timer
    start_time = time.time()
    for i, element in enumerate(elements):
        year_span = span_elements[i] if i < len(span_elements) else None
        year_of_publication = year_span.text if year_span else "N/A"
        
        cite_span = cite_elements[i] if i < len(cite_elements) else None
        no_of_title_cites = cite_span.text if cite_span else "N/A"

        title = element.text
        link = element.get_attribute('href')

        # Use PaperScraper to get additional details
        details = paper_scraper.scrape_paper_details(link)
        print(details)

        
        # Renew the Tor connection every 10 seconds
        if time.time() - start_time > 30:
          print("Renewing Tor connection...")
          tor_proxy.renew_connection()
          proxy_ip = tor_proxy.get_ip()
          print(f'previous IP: {proxy_ip}')
          
          # Reset the timer
          start_time = time.time()
          
        
        paper_detail = {
            'NAME': hyperlink[0],
            'TITLE': title,
            'YEAR': year_of_publication,
            'LINK': link,
            'CITATIONS': citation_metrics['CITATIONS'],
            'H_INDEX': citation_metrics['H_INDEX'],
            'I10_INDEX': citation_metrics['I10_INDEX'],
            'TITLE_CITES': no_of_title_cites,
            'AUTHORS': details['authors'],
            'JOURNAL': details['journal'],
            'VOLUME': details['volume'],
            'PAGES': details['pages'],
            'BOOKTITLE': details['booktitle'],
            'ORGANIZATION': details['organization']
        }

        paper_details.append(paper_detail)

driver.quit()
paper_scraper.close()

# Save the paper details to a CSV file
csv_file = "research_papers.csv"
fieldnames = ['NAME', 'TITLE', 'LINK', 'YEAR', 'CITATIONS', 'H_INDEX', 'I10_INDEX', 'TITLE_CITES', 
              'AUTHORS', 'JOURNAL', 'VOLUME', 'PAGES', 'BOOKTITLE', 'ORGANIZATION']

with open(csv_file, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.DictWriter(file, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(paper_details)

print(f"Research paper details saved to {csv_file}")
