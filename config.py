from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

service = Service()
options=Options()
#options.add_argument("--headless=new")

driver = webdriver.Chrome(service=service, options=options)
