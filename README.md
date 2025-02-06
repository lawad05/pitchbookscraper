# Pitchbook Scraper README
**Overview**
This Python script automates the process of extracting company data from PitchBook using Selenium and BeautifulSoup. The extracted data is saved in an Excel file formatted with OpenPyXL. This tool is designed for finance professionals who need to collect and organize company information efficiently.    
**Features**
- Searches for multiple companies on PitchBook.
- Extracts key details, including:
- Company name, entity type, website, headquarters, and industries.
- Primary contact details, overview, financials, products/services, and recent news.
- Exports data to a neatly formatted Excel file.
- Automatically adjusts Excel column widths and adds a logo to the report.

**Requirements**
- Python 3.7+
- Google Chrome with debugging enabled (--remote-debugging-port=9222).
- Dependencies:
pip install selenium beautifulsoup4 pandas openpyxl

**Setup Instructions**
1. Enable Chrome Debugging Mode: Run Chrome with the following command to allow Selenium to connect to an existing session:
google-chrome --remote-debugging-port=9222
2. Edit Configuration:
- Update the path to your logo image in the save_to_excel() function:
logo_path = "/path/to/your/logo.jpg"
3. Install the required packages:
pip install selenium beautifulsoup4 pandas openpyxl
4. Ensure the Chrome WebDriver is installed and in your PATH.

**Usage Instructions**
1. Run the script:
python pitchbook_scraper.py
2. Input the Required Information:
Enter the company names you want to search, separated by commas.
Specify the name for the Excel file.
3. The script will:
Search for each company on PitchBook.
Extract and display relevant data.
Save the information in an Excel file with formatting and a company logo.

**Notes and Recommendations**
Error Handling: The script includes try-except blocks to handle various errors, such as missing elements or timeouts.
Performance: Depending on the number of companies, the script may take time to navigate PitchBook and extract information.
Selenium Chrome Setup: Ensure you have a stable Chrome WebDriver compatible with your Chrome version.  

**Troubleshooting**
If the script cannot find elements on PitchBook, verify that the site structure hasn't changed.
Ensure Chrome is running in debugging mode (--remote-debugging-port=9222).
Check that your WebDriver is up-to-date.  
