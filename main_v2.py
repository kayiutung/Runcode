#Import neccessary packages
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from datetime import datetime, date, timedelta
import pandas as pd
import time

#Set webdriver
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

#Store all scrapped data in a list
def store_sale(data):  # Store scrapped sale data to the list as int
  if "," in data:
    data = data.replace(",","")
  else:
    pass
  Sale.append(int(data))

def store_lease(data):  # Store scrapped lease data to the list as int
  if "," in data:
    data = data.replace(",","")
  else:
    pass
  Lease.append(int(data))

#Create lists
Sale = []
Lease = []
d = date.today() + timedelta(days=1)
print(d)
Sale.append(d)
Lease.append(d)

#Read existing data
df_sale = pd.read_excel("中原放盤.xlsx", sheet_name="賣盤")
df_lease = pd.read_excel("中原放盤.xlsx", sheet_name="租盤")

#Sale
driver = webdriver.Chrome('/usr/bin/chromedriver', options = options)
url = "https://hk.centanet.com/findproperty/list/buy"
driver.get(url)

data = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[6]/div/div[1]/div[1]/div/h2/span/span').text
print(data)
store_sale(data)
time.sleep(3)

for price in range(100,2001,100):
    driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[3]/div/div[2]/div[1]/div[4]/span/button/i').click()
    time.sleep(3)
    driver.find_element_by_xpath('/html/body/ul/div[2]/div[2]/div[1]/div/input').send_keys(Keys.BACKSPACE + Keys.BACKSPACE + Keys.BACKSPACE + Keys.BACKSPACE + Keys.BACKSPACE + Keys.BACKSPACE)
    driver.find_element_by_xpath('/html/body/ul/div[2]/div[2]/div[1]/div/input').send_keys(price-99)
    time.sleep(3)
    driver.find_element_by_xpath('/html/body/ul/div[2]/div[2]/div[2]/div/input').send_keys(Keys.BACKSPACE + Keys.BACKSPACE + Keys.BACKSPACE + Keys.BACKSPACE + Keys.BACKSPACE + Keys.BACKSPACE)
    driver.find_element_by_xpath('/html/body/ul/div[2]/div[2]/div[2]/div/input').send_keys(price)
    time.sleep(3)
    driver.find_element_by_xpath('/html/body/ul/div[2]/div[2]/div[2]/div/input').send_keys(Keys.ENTER)
    time.sleep(3)

    data = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[6]/div/div[1]/div[1]/div/h2/span/span').text
    print(data)
    store_sale(data)

driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[3]/div/div[2]/div[1]/div[4]/span/button/i').click()
time.sleep(3)
driver.find_element_by_xpath('/html/body/ul/div[2]/div[2]/div[1]/div/input').clear()
driver.find_element_by_xpath('/html/body/ul/div[2]/div[2]/div[1]/div/input').send_keys(2001)
time.sleep(3)
driver.find_element_by_xpath('/html/body/ul/div[2]/div[2]/div[1]/div/input').send_keys(Keys.ENTER)
time.sleep(3)
data = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[6]/div/div[1]/div[1]/div/h2/span/span').text
print(data)
store_sale(data)

driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[3]/div/div[2]/div[1]/div[4]/span/button/i[1]').click()
time.sleep(3)

for size in range(200,1001,100):
    driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[3]/div/div[2]/div[1]/div[5]/span/button/i').click()
    time.sleep(3)
    driver.find_element_by_xpath('/html/body/ul[2]/div[2]/div[2]/div[1]/div/input').send_keys(Keys.BACKSPACE + Keys.BACKSPACE + Keys.BACKSPACE + Keys.BACKSPACE)
    driver.find_element_by_xpath('/html/body/ul[2]/div[2]/div[2]/div[1]/div/input').send_keys(size-99)
    time.sleep(3)
    driver.find_element_by_xpath('/html/body/ul[2]/div[2]/div[2]/div[2]/div/input').send_keys(Keys.BACKSPACE + Keys.BACKSPACE + Keys.BACKSPACE + Keys.BACKSPACE)
    driver.find_element_by_xpath('/html/body/ul[2]/div[2]/div[2]/div[2]/div/input').send_keys(size)
    time.sleep(3)
    driver.find_element_by_xpath('/html/body/ul[2]/div[2]/div[2]/div[2]/div/input').send_keys(Keys.ENTER)
    time.sleep(3)

    data = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[6]/div/div[1]/div[1]/div/h2/span/span').text
    print(data)
    store_sale(data)

driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[3]/div/div[2]/div[1]/div[5]/span/button/i').click()
time.sleep(3)
driver.find_element_by_xpath('/html/body/ul[2]/div[2]/div[2]/div[1]/div/input').clear()
driver.find_element_by_xpath('/html/body/ul[2]/div[2]/div[2]/div[1]/div/input').send_keys(1001)
time.sleep(3)
driver.find_element_by_xpath('/html/body/ul[2]/div[2]/div[2]/div[1]/div/input').send_keys(Keys.ENTER)
time.sleep(3)
data = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[6]/div/div[1]/div[1]/div/h2/span/span').text
print(data)
store_sale(data)

driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[3]/div/div[2]/div[1]/div[5]/span/button/i[1]').click()
time.sleep(3)

driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[3]/div/div[2]/div[1]/div[6]/span/button/i').click()
time.sleep(3)

for fp in range(1,6):
    driver.find_element_by_xpath('/html/body/ul[3]/div[2]/div/label[' + str(fp) + ']/span[1]/span').click()
    time.sleep(3)
    data = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[6]/div/div[1]/div[1]/div/h2/span/span').text
    print(data)
    store_sale(data)

    driver.find_element_by_xpath('/html/body/ul[3]/div[2]/div/label[' + str(fp) + ']/span[1]/span').click()
    time.sleep(3)

driver.quit()

#Lease
driver = webdriver.Chrome('/usr/bin/chromedriver', options = options)
url = "https://hk.centanet.com/findproperty/list/rent"
driver.get(url)

data = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[6]/div/div[1]/div[1]/div/h2/span/span').text
print(data)
store_lease(data)
time.sleep(3)

for size in range(200,1001,100):
    driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[3]/div/div[2]/div[1]/div[5]/span/button/i').click()
    time.sleep(3)
    driver.find_element_by_xpath('/html/body/ul/div[2]/div[2]/div[1]/div/input').send_keys(Keys.BACKSPACE + Keys.BACKSPACE + Keys.BACKSPACE + Keys.BACKSPACE)
    driver.find_element_by_xpath('/html/body/ul/div[2]/div[2]/div[1]/div/input').send_keys(size-99)
    time.sleep(3)
    driver.find_element_by_xpath('/html/body/ul/div[2]/div[2]/div[2]/div/input').send_keys(Keys.BACKSPACE + Keys.BACKSPACE + Keys.BACKSPACE + Keys.BACKSPACE)
    driver.find_element_by_xpath('/html/body/ul/div[2]/div[2]/div[2]/div/input').send_keys(size)
    time.sleep(3)
    driver.find_element_by_xpath('/html/body/ul/div[2]/div[2]/div[2]/div/input').send_keys(Keys.ENTER)
    time.sleep(3)

    data = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[6]/div/div[1]/div[1]/div/h2/span/span').text
    print(data)
    store_lease(data)

driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[3]/div/div[2]/div[1]/div[5]/span/button/i').click()
time.sleep(3)
driver.find_element_by_xpath('/html/body/ul/div[2]/div[2]/div[1]/div/input').clear()
driver.find_element_by_xpath('/html/body/ul/div[2]/div[2]/div[1]/div/input').send_keys(1001)
time.sleep(3)
driver.find_element_by_xpath('/html/body/ul/div[2]/div[2]/div[1]/div/input').send_keys(Keys.ENTER)
time.sleep(3)
data = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[6]/div/div[1]/div[1]/div/h2/span/span').text
print(data)
store_lease(data)

driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[3]/div/div[2]/div[1]/div[5]/span/button/i[1]').click()
time.sleep(3)

driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[3]/div/div[2]/div[1]/div[6]/span/button/i').click()
time.sleep(3)

for fp in range(1,6):
    driver.find_element_by_xpath('/html/body/ul[2]/div[2]/div/label[' + str(fp) + ']/span[1]/span').click()
    time.sleep(3)
    data = driver.find_element_by_xpath('//*[@id="__layout"]/div/div[4]/div[6]/div/div[1]/div[1]/div/h2/span/span').text
    print(data)
    store_lease(data)

    driver.find_element_by_xpath('/html/body/ul[2]/div[2]/div/label[' + str(fp) + ']/span[1]/span').click()
    time.sleep(3)

driver.quit()

# Add the scrapped data to df and format the datetime
df_sale.loc[len(df_sale)] = Sale
df_lease.loc[len(df_lease)] = Lease
df_sale['日期'] = pd.to_datetime(df_sale['日期'])
df_sale['日期'] = df_sale['日期'].dt.strftime('%Y-%m-%d')
df_lease['日期'] = pd.to_datetime(df_lease['日期'])
df_lease['日期'] = df_lease['日期'].dt.strftime('%Y-%m-%d')

# Put the df back to the Excel
with pd.ExcelWriter("中原放盤.xlsx") as writer:
  df_sale.to_excel(writer, sheet_name="賣盤", index=False)
  df_lease.to_excel(writer, sheet_name="租盤", index=False)
