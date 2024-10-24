# Web-Scraper-and-Sentiment-Analysis
## Overview
This Python script performs web scraping and sentiment analysis of web pages provided in an Excel file (Input.xlsx). The script extracts textual content from web pages, preprocesses the text, and calculates various sentiment and readability metrics. It saves the results in an output Excel file (OutputDataStructure.xlsx).

## Prerequisites
To run this script, the following Python libraries need to be installed:

openpyxl: To read and write Excel files.
requests: To send HTTP requests and retrieve web page content.
beautifulsoup4: To parse the HTML content of web pages.
You can install these dependencies using pip:

```
pip install openpyxl requests beautifulsoup4
```
## Features
1. Web Scraping:

- The script reads URLs from the Input.xlsx file and fetches the web content using the requests library.
- The title and content of the page are extracted using BeautifulSoup and saved to text files named based on the URL IDs.

2. Preprocessing:

- The textual content is converted to lowercase and cleaned by removing punctuation.
- The content is tokenized into individual words for further analysis.

3. Sentiment Analysis:

The script computes sentiment scores using a predefined list of positive and negative words.
- Metrics calculated:
- Positive and negative scores
- Polarity score
- Subjectivity score
- Readability Analysis:

The script also evaluates the readability of the web content by calculating:
  - Average sentence length
  - Percentage of complex words
  - Fog index
  - Word count and average word length
  - Number of personal pronouns
  
- Excel Output:
  - The calculated sentiment and readability metrics are saved into the OutputDataStructure.xlsx file.

## How to Use
1. Prepare the Input File:

Create an Input.xlsx file with two columns:
-URL ID (Unique Identifier)
-URL (Web page link)

2. Run the Script:

- Ensure that the Input.xlsx file is in the same directory as the script or modify the path to it in the script.
- Execute the script in your Python environment:
```
python main.py
```

3. View the Results:

- The scraped web content will be saved as text files.
- The sentiment and readability scores will be stored in the OutputDataStructure.xlsx file.

## Configuration
You may need to modify the paths in the script:
- Ensure that the paths for reading the Input.xlsx file and saving the OutputDataStructure.xlsx file match the locations on your system.

## Notes
- The script assumes that the HTML structure of the web pages follows a certain format. If the pages have a different structure, the extraction logic may need to be updated.
