import os
import psutil
import time
import random
import logging
import pyperclip as pc
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
from fp.fp import FreeProxy

class ScholarCitationFetcher:
    def __init__(self):
        self.logger = self._setup_logger()
        self.proxies = self._get_proxies()
        self.driver = self._setup_driver()

    def _setup_logger(self):
        logging.basicConfig(level=logging.INFO)
        return logging.getLogger(__name__)

    def _get_proxies(self):
        # Use FreeProxy to fetch proxies
        fp = FreeProxy()
        proxies = fp.get()
        if not proxies:
            self.logger.warning("No proxies found.")
        else:
            self.logger.info(f"Fetched {len(proxies)} proxies.")
        return proxies

    def _setup_driver(self):
        self._close_chrome_instances()
        options = self._configure_chrome_options()
        return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    def _close_chrome_instances(self):
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] == 'chrome.exe':
                proc.kill()
        self.logger.info("All Chrome instances have been closed.")

    def _configure_chrome_options(self):
        options = Options()
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--start-maximized')
        # Add these lines to use your specific user profile
        options.add_argument('--remote-debugging-port=9222')
        options.add_argument('--user-data-dir=C:\\Users\\User\\AppData\\Local\\Google\\Chrome\\User Data')
        options.add_argument('--profile-directory=Profile 9')
        #options.add_argument('--headless')
        chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
        if not os.path.exists(chrome_path):
            self.logger.error(f"Chrome executable not found at {chrome_path}")
            exit(1)
        options.binary_location = chrome_path

        # Set a random proxy
        if self.proxies:
            proxy = random.choice(self.proxies)
            options.add_argument(f'--proxy-server=http://{proxy}')
            self.logger.info(f"Using proxy: {proxy}")

        return options

    def fetch_citation(self, paper_title, year):
        try:
            base_url = f"https://scholar.google.com/scholar?hl=en&q={paper_title} {year}"
            self.driver.get(base_url)
            time.sleep(random.uniform(2, 5))

            block = self.driver.find_element(By.CSS_SELECTOR, "div.gs_ri")
            if not block:
                self.logger.warning("No results found.")
                return None

            title = block.find_element(By.CSS_SELECTOR, "h3")
            link = title.find_element(By.TAG_NAME, "a") if title else None
            if not link:
                self.logger.warning("Title link not found.")
                return None

            citation_id = link.get_attribute("data-clk-atid")
            if not citation_id:
                self.logger.warning("Citation ID not found.")
                return None

            cite_url = f"https://scholar.google.com/scholar?hl=en&q=info:{citation_id}:scholar.google.com/&output=cite&scirp=0"
            self.driver.get(cite_url)
            time.sleep(random.uniform(2, 5))

            latex_link_tag = self.driver.find_element(By.LINK_TEXT, "BibTeX")
            if not latex_link_tag:
                self.logger.warning("BibTeX link not found.")
                return None

            latex_link = latex_link_tag.get_attribute("href")
            if not latex_link:
                self.logger.warning("BibTeX link href not found.")
                return None

            self.driver.get(latex_link)
            wait = WebDriverWait(self.driver, 10)
            pre_element = wait.until(EC.presence_of_element_located((By.TAG_NAME, "pre")))
            citation = pre_element.text

            return self._parse_bibtex(citation)

        except Exception as e:
            self.logger.error(f"An error occurred: {str(e)}")
            return None

        finally:
            self.driver.quit()

    def _parse_bibtex(self, bibtex):
        data = {
            'AUTHORS': self._extract_field(bibtex, 'author'),
            'JOURNAL': self._extract_field(bibtex, 'journal') or self._extract_field(bibtex, 'booktitle') or 'N/A',
            'VOLUME': self._extract_field(bibtex, 'volume') or 'N/A',
            'PAGES': self._extract_field(bibtex, 'pages') or 'N/A',
            'PUBLISHER': self._extract_field(bibtex, 'publisher') or 'N/A'
        }
        return data

    def _extract_field(self, bibtex, field):
        match = re.search(rf'{field}\s*=\s*{{(.*?)}}', bibtex, re.IGNORECASE | re.DOTALL)
        return match.group(1).strip() if match else None

if __name__ == "__main__":
    fetcher = ScholarCitationFetcher()
    result = fetcher.fetch_citation("Autonomous Electromagnetic Signal Analysis and Measurement System", "2024")
    print(result)

    
    
    
