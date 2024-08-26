import requests
from bs4 import BeautifulSoup
import pyperclip as pc

# Pre-set paper title
paper = "Autonomous Electromagnetic Signal Analysis and Measurement System"

# Construct the Google Scholar search URL
base_url = f"https://scholar.google.com/scholar?hl=en&q={paper}"
print(base_url)

# Send a GET request to Google Scholar
response = requests.get(base_url)

# Parse the response content with BeautifulSoup
soup = BeautifulSoup(response.content, "html.parser")

# Find the first search result block
block = soup.find("div", {"class": "gs_ri"})

# Extract the title and link
title = block.find("h3")
link = title.find("a")

# Get the citation ID from the link
citation_id = link["data-cid"]

# Construct the citation URL
cite_url = f"https://scholar.google.com/scholar?q=info:{citation_id}:scholar.google.com/&output=citation&scisig=AAGBfm0AAAAAZrHC7FrQRY7jF1accZg5V_rYLTw"

# Send a GET request to the citation URL
cite_response = requests.get(cite_url)

# Parse the citation page content with BeautifulSoup
cite_soup = BeautifulSoup(cite_response.content, "html.parser")

# Find the LaTeX citation link
latex_link = cite_soup.find("a", text="BibTeX")["href"]

# Send a GET request to the LaTeX citation link
latex_response = requests.get(latex_link)

# Extract the citation text
citation = latex_response.text

# Copy the citation text to the clipboard
pc.copy(citation)

# Print the citation text
print(citation)
