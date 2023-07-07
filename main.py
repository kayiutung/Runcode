from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd
from datetime import date, datetime, timedelta


#set webdriver
chrome_options = webdriver.chrome.options.Options()

prefs = {
'download.extensions_to_open': 'xml',
'safebrowsing.enabled': True
}
## Uncomment the below to let the program to run in background
options = webdriver.ChromeOptions()
options.add_experimental_option('prefs',prefs)
options.add_argument("start-maximized")
options.add_argument("--headless")
# options.add_argument("disable-infobars")
options.add_argument("--disable-extensions")
options.add_argument("--safebrowsing-disable-download-protection")
options.add_argument("safebrowsing-disable-extension-blacklist")
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')

Date = [date(2023, 7, 1), date(2023, 7, 2)]
Sale = [39762, 39756]
Lease = [14068, 14084]

#Sale
driver = webdriver.Chrome('/home/runner/work/chromedriver', options = options)
url = 'https://hk.centanet.com/findproperty/list/buy'
driver.get(url)
sale = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[6]/div/div[1]/div[1]/div/h2/span/span').text
print(sale)
Sale.append(sale)
driver.quit()

#Lease
driver = webdriver.Chrome('/home/runner/work/chromedriver', options = options)
url = 'https://hk.centanet.com/findproperty/list/rent'
driver.get(url)
lease = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[6]/div/div[1]/div[1]/div/h2/span/span').text
print(lease)
Lease.append(lease)
driver.quit()

#Date
d = date.today() + timedelta(days=1)
print(d)
Date.append(d)

df=pd.DataFrame()
df["日期"] = Date
df["賣盤"] = Sale
df["租盤"] = Lease
df.to_excel("Centaline_sale_lease.xlsx",sheet_name="sheet1", index=False)
