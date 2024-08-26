import re
import csv
import time
import logging
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException
from tor_proxy import TorProxy  # Import the TorProxy class

class PaperScraper:
    def __init__(self):
        # Set up logging
        logging.basicConfig(filename='paper_scraper.log', level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
        
        # Initialize the TorProxy
        self.proxy = TorProxy()
        self.proxy.start()
        logging.info("TorProxy started")

        options = Options()
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        # Add these lines to use your specific user profile
        options.add_argument('--user-data-dir=C:\\Users\\User\\AppData\\Local\\Google\\Chrome\\User Data')
        options.add_argument('--profile-directory=Profile 9')
        #options.add_argument('--headless')  

        # Set proxy for Chrome
        proxy_address = "socks5://localhost:9055"
        options.add_argument(f'--proxy-server={proxy_address}')
        options.add_argument('--host-resolver-rules="MAP * ~NOTFOUND, EXCLUDE localhost"')
        
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        logging.info("Webdriver initialized")

    def process_authors(self, match):
        first_group = match.group(1)[0]
        if match.group(3):
            second_group = match.group(2)[0]
            return f"{match.group(3)} {first_group}. {second_group}."
        else:
            return f"{match.group(2)} {first_group}."

    def scrape_paper_details(self, url):
        self.driver.get(url)
        time.sleep(1)

        details = {
            'authors': 'N/A',
            'journal': 'N/A',
            'volume': 'N/A',
            'pages': 'N/A',
            'booktitle': 'N/A',
            'organization': 'N/A'
        }

        # Authors
        try:
            authors_field = self.driver.find_element(By.XPATH, '//div[@class="gsc_oci_field" and contains(text(), "Authors")]')
            authors_value = authors_field.find_element(By.XPATH, './following-sibling::div[@class="gsc_oci_value"]')
            authors_text = authors_value.text
            details['authors'] = re.sub(r'(\w+)\s+(\w+)(?:\s+(\w+))?', self.process_authors, authors_text)
        except:
            try:
                authors_field = self.driver.find_element(By.XPATH, '//div[@class="gsc_oci_field" and contains(text(), "Inventors")]')
                authors_value = authors_field.find_element(By.XPATH, './following-sibling::div[@class="gsc_oci_value"]')
                authors_text = authors_value.text
                details['authors'] = re.sub(r'(\w+)\s+(\w+)(?:\s+(\w+))?', self.process_authors, authors_text)
            except:
                logging.warning(f"Failed to extract authors for {url}")

        # Journal/Book/Source
        for field in ["Journal", "Book", "Source"]:
            try:
                journal_field = self.driver.find_element(By.XPATH, f'//div[@class="gsc_oci_field" and contains(text(), "{field}")]')
                journal_value = journal_field.find_element(By.XPATH, './following-sibling::div[@class="gsc_oci_value"]')
                details['journal'] = journal_value.text
                break
            except:
                pass
        
        if details['journal'] == 'N/A':
            logging.warning(f"Failed to extract journal/book/source for {url}")

        # Volume
        try:
            volume_field = self.driver.find_element(By.XPATH, '//div[@class="gsc_oci_field" and contains(text(), "Volume")]')
            volume_value = volume_field.find_element(By.XPATH, './following-sibling::div[@class="gsc_oci_value"]')
            details['volume'] = volume_value.text
        except:
            logging.warning(f"Failed to extract volume for {url}")

        # Pages
        try:
            pages_field = self.driver.find_element(By.XPATH, '//div[@class="gsc_oci_field" and contains(text(), "Pages")]')
            pages_value = pages_field.find_element(By.XPATH, './following-sibling::div[@class="gsc_oci_value"]')
            details['pages'] = pages_value.text
        except:
            logging.warning(f"Failed to extract pages for {url}")

        # Booktitle (for conference papers)
        try:
            booktitle_field = self.driver.find_element(By.XPATH, '//div[@class="gsc_oci_field" and contains(text(), "Conference")]')
            booktitle_value = booktitle_field.find_element(By.XPATH, './following-sibling::div[@class="gsc_oci_value"]')
            details['booktitle'] = booktitle_value.text
        except:
            logging.warning(f"Failed to extract booktitle for {url}")

        # Organization (for conference papers)
        try:
            org_field = self.driver.find_element(By.XPATH, '//div[@class="gsc_oci_field" and contains(text(), "Publisher")]')
            org_value = org_field.find_element(By.XPATH, './following-sibling::div[@class="gsc_oci_value"]')
            details['organization'] = org_value.text
        except:
            logging.warning(f"Failed to extract organization for {url}")
 
        return details

    def is_detected(self):
        detected = False
        
        # Check for captcha
        try:
            if self.driver.find_element(By.ID, "captcha-form"):
                logging.warning("Captcha detected")
                detected = True
        except:
            pass
        
        # Check for unusual traffic message
        try:
            if self.driver.find_element(By.XPATH, "//div[contains(text(), 'unusual traffic')]"):
                logging.warning("Unusual traffic message detected")
                detected = True
        except:
            pass
        
        if detected:
            logging.info("Detection triggered. Renewing Tor connection...")
            self.proxy.renew_connection()
            time.sleep(10)  # Wait for new connection to establish
            return True
        
        return False

    def handle_network_issue(self):
        logging.warning("Network issue detected. Attempting to resolve...")
        
        # Renew Tor connection
        self.proxy.renew_connection()
        time.sleep(10)  # Wait for new connection to establish
        
        # Check if the connection is restored
        try:
            self.driver.get("https://www.google.com")
            if "google.com" in self.driver.current_url:
                logging.info("Network connection restored")
                return True
        except WebDriverException:
            logging.error("Failed to restore network connection")
            return False

    def scrape_and_parse(self, input_file, output_file, renew_interval=5, max_retries=7):
        with open(input_file, mode='r', encoding='utf-8') as file, \
             open(output_file, mode='w', newline='', encoding='utf-8') as csv_file:
            reader = csv.DictReader(file)
            fieldnames = ['NAME', 'AUTHORS', 'YEAR', 'TITLE', 'JOURNAL', 'VOLUME', 'PAGES', 'BOOKTITLE', 'ORGANIZATION']
            writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
            writer.writeheader()

            for count, row in enumerate(reader, start=1):
                link = row['LINK']
                name = row['NAME']
                title = row.get('TITLE', 'N/A')
                year = row.get('YEAR', 'N/A')
                logging.info(f"Scraping: {name}")

                retries = 0
                while retries < max_retries:
                    try:
                        details = self.scrape_paper_details(link)
                        
                        if self.is_detected():
                            raise Exception("Google detection triggered")
                        
                        writer.writerow({
                            'NAME': name,
                            'AUTHORS': details['authors'],
                            'YEAR': year,
                            'TITLE': title,
                            'JOURNAL': details['journal'],
                            'VOLUME': details['volume'],
                            'PAGES': details['pages'],
                            'BOOKTITLE': details['booktitle'],
                            'ORGANIZATION': details['organization']
                        })
                        logging.info(f"Successfully scraped: {name}")
                        break  # Successfully scraped, exit the retry loop
                    except WebDriverException as e:
                        logging.error(f"Network error while scraping {name}: {str(e)}")
                        if not self.handle_network_issue():
                            logging.error("Unable to resolve network issue. Skipping this paper.")
                            break
                        retries += 1
                    except Exception as e:
                        logging.error(f"Error scraping {name}: {str(e)}")
                        retries += 1
                        if retries < max_retries:
                            logging.info(f"Retrying... (Attempt {retries + 1} of {max_retries})")
                            self.proxy.renew_connection()
                            time.sleep(10)  # Wait for new connection to establish
                        else:
                            logging.error(f"Failed to scrape {name} after {max_retries} attempts")
                            writer.writerow({
                                'NAME': name,
                                'AUTHORS': 'N/A',
                                'YEAR': year,
                                'TITLE': title,
                                'JOURNAL': 'N/A',
                                'VOLUME': 'N/A',
                                'PAGES': 'N/A',
                                'BOOKTITLE': 'N/A',
                                'ORGANIZATION': 'N/A'
                            })

                # Renew Tor connection every `renew_interval` papers
                if count % renew_interval == 0:
                    logging.info("Renewing Tor connection...")
                    self.proxy.renew_connection()
                    time.sleep(10)  # Give Tor some time to establish a new connection

        logging.info(f"SAVED TO {output_file}")

    def close(self):
        self.driver.quit()
        self.proxy.stop()  # Stop Tor when done
        logging.info("Scraper closed and Tor proxy stopped")

# Example usage:
# scraper = PaperScraper()
# scraper.scrape_and_parse('research_papers.csv', 'Research_paper_details.csv')
# scraper.close()