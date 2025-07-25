🏦 Stuttgart Bond Search Automation
Last major update: 23.4.2025
NOTE! At the the moment (25.7.2025) this version doesn't work anymore, due to change in structure of the scraped page!!

🎉 Major refactor! Cookie consent handling now uses <aside> elements. Many improvements – essentially, everything changed.

📌 Overview
This Python project automates the process of searching for bonds on the Börse Stuttgart website using Selenium. It selects specific criteria, interacts with dynamic filters, extracts the resulting data, and writes it to an Excel file. You can run the script periodically via crontab in a headless mode.

At the end, a GPT-4 summary is optionally generated to provide insights about the bonds (requires OpenAI API access).

🧩 Project Structure

├── bond_search5.py            # Main script

├── web_actions.py           # Utility functions for Selenium interactions

├── web_actions.py           # Utility functions for Selenium interactions

├── config_webdriver_manager.py # Crome web driver setup

├── locators.py              # Central place for XPath and CSS selectors

├── gpt_helper.py            # Interacts with GPT-4 for summaries (optional)

├── requirements.txt         # Required Python packages

⚙️ Setup Instructions

🔐 Prerequisites
Python 3.8+

Chrome browser

Chromedriver (Windows only; handled automatically in Linux via webdriver_manager)

🐍 Create and activate virtual environment (Linux example)
bash
Copier
Modifier
python3 -m venv venv
source venv/bin/activate
📦 Install dependencies
bash
Copier
Modifier
pip install -r requirements.txt
Note: Ensure your requirements.txt includes packages such as selenium, pandas, numpy, openpyxl, bs4, webdriver_manager, etc.

🧠 Key Features
Automated bond search with custom filters:

Currency (Euro)

Tradable units

Listed date range

Maturity date range

Bond type (Unternehmensanleihe)

Minimum yield (e.g., 3%)

Scrolls and clicks "show more results"

Extracts table data using BeautifulSoup

Exports results to Excel with formatted output

Automatically names Excel files with the current date

Optional summary using GPT-4:

Highlights the highest yield bond

Lists common maturity years

Notes interesting patterns

🤖 Headless Mode
To run in headless mode (for crontab etc.), uncomment the relevant line in config.py:

python
Copier
Modifier
chrome_options.add_argument("--headless")
📊 Output
Excel file saved to c:/temp/tutorial/sample.xlsx (adjust path as needed)

Optional GPT-4 text summary saved to c:/temp/tutorial/bond_summary.txt

📅 Automation with Crontab (Linux)
Example crontab entry to run daily at 10 AM:

bash
Copier
Modifier
0 10 * * * /path/to/venv/bin/python /path/to/bond_search/bond_search.py
Make sure the script has execution permissions and proper download paths.

🧪 Example Locators Structure
python
Copier
Modifier
class Locators:
    cookie_acceptor = ".js-bsg-cookie-layer__confirm > span:nth-child(1)"
    Cookie_einstellungen_XPATH = '//*[@id="cookie-consent"]/div[2]/div[3]/button[1]/span'
    ...
Organizing all selectors centrally improves maintainability as the page structure changes.

🧠 GPT Integration
To use GPT-4 for bond analysis, you need to:

Set your OpenAI API key in an environment variable or config file

Ensure gpt_helper.py contains your API call logic

✅ To-Do / Ideas
Add logging instead of print() statements

Support multiple export formats (CSV, JSON)

Add CLI arguments for flexible date filtering

Dockerize the project for easier deployment

📝 License
MIT License. Use and modify freely, but no guarantees or warranties.

