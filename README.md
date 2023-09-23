# Selenium-Stuttgart-Boerse
A Python program with  Selenium module, to open "Stuttgart" Bond search, select criteria and output the result to a excel file.
There are some dates calculated from the now with the datetime, and years are added or subtracted and the calculated dates are entered in the web page. THe excel file name contains the actural date. My idea was to run this program in certain intervals, automatically started y crontab, and then it would be "headless".

Bond_search.py is the main program, and to have more structure, 3 different modules are imported: config, web actions and locators

1. config.py is the webdriver section in linux. the config1.py works for me in Windows. In Linux, I didn't have to download any Chromedriver, but in Windows version it was necessary.  

2. web_actions.py contains different functions, depending on whether id or xpath etc is used

3. locators.py contains the "library" of different xpath or id's with a description of the field, as the XPath or id is not descriptive

I used the program in a virtual environment. 

for installing the requirements.txt enter in terminal:
pip install -r requirements.text


Note to myself: the virtual environment i created by entering in terminal:
python3 -m venv venv

to start the virtual environment enter in terminal:
source venv/bin/activate
