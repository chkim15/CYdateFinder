import pandas as pd
from datetime import date, datetime, timedelta
from dateutil.relativedelta import relativedelta

from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

import time
import os

# don't show chrome browser
options = Options()
options.add_argument("--headless")

if  getattr(sys, 'frozen', False): 
    chromedriver_path = os.path.join(sys._MEIPASS, "chromedriver.exe")
    driver = webdriver.Chrome(chromedriver_path, options=options)
else:
    driver = webdriver.Chrome(options=options)


################### BELOW IS FOR DESTINATION: DUBAI, DAMMAM, RIYADH #############################
# open ONE shipper's website
driver.get("https://ecomm.one-line.com/one-ecom")
time.sleep(3)

# close initial pop-up
driver.find_element(By.CLASS_NAME,"PrePromotionPopup_close-icon__WOnwe").click()

# close cookie policy pop-up at the bottom
driver.find_element(By.XPATH,"//button[@class='CookiePolicy_closeButton__6sDK1']").click()

# type in origin (input-based dropdown)
origin = driver.find_element(By.ID,"downshift-1-input")
origin.send_keys("PUSAN")
time.sleep(1)
origin.send_keys(Keys.DOWN)
origin.send_keys(Keys.RETURN)

# type in destination (input-based dropdown) Dubai = Jebel Ali
destination = driver.find_element(By.ID,"downshift-2-input")
destination.send_keys("JEBEL ALI")
time.sleep(1)
destination.send_keys(Keys.DOWN)
destination.send_keys(Keys.RETURN)

# next (how many weeks to search for)
driver.find_element(By.XPATH,"//div[@class='Input_inputContainer__TLu7L']/input[@data-cy='port-next-value']").click()

# select next 8 weeks
driver.find_element(By.ID,"next-week-value-56").click()

# click 'Search'
driver.find_element(By.ID,"schedule-box-search-btn-id").click()
time.sleep(2)

# get departure dates
depart_dates_get = driver.find_elements(By.XPATH,"//div[@class='ScheduleItem_date__WLAPW'][@data-cy='new-schedule-departure']")

depart_dates = []
for date in depart_dates_get:
    depart_dates.append(date.text)

# get arrival dates
arrival_dates_get = driver.find_elements(By.XPATH,"//div[@class='ScheduleItem_date__WLAPW'][@data-cy='new-schedule-arrival']")

arrival_dates = []
for date in arrival_dates_get:
    arrival_dates.append(date.text)

# get vessel names
vessel_names_get = driver.find_elements(By.XPATH,"//a[@data-cy='new-schedule-vessel-name']")

vessel_names = []
for vessel in vessel_names_get:
    vessel_names.append(' '.join(vessel.text.split()[:-1]))

# get service lane info
service_lanes_get = driver.find_elements(By.XPATH,"//a[@data-cy='new-schedule-service-land-name']")

service_lanes = []
for lane in service_lanes_get:
    service_lanes.append(lane.text)

# create dataframe for Dubai port (based on ONE shipper's website)
df_shipper_dubai = pd.DataFrame({'departure': depart_dates, 'arrival': arrival_dates, 'vessel': vessel_names, 'route': service_lanes})

# open Busan port terminal website
driver.get("https://www.hpnt.co.kr/infoservice/vessel/vslScheduleList.jsp")

from datetime import date as ddt
# get dates for two months from today
today = ddt.today()
after_2months = str(today + relativedelta(months=+2))

# change ending date for search
end_date = driver.find_element(By.ID,"strdEdDate")
end_date.clear()
end_date.send_keys(after_2months)

# type in route and click 'search'
route = driver.find_element(By.ID,"route")
route.send_keys("AG3")
driver.find_element(By.ID,"submitbtn").click()
time.sleep(1)

# get vessel's name
vessel_terminal_get = driver.find_elements(By.XPATH,"//tr[@class='color_planned']/td[5]")
vessel_terminal = []
for vessel in vessel_terminal_get:
    vessel_terminal.append(vessel.text)

# get expected anchoring date
anchor_get = driver.find_elements(By.XPATH,"//tr[@class='color_planned']/td[8]")
anchor_dates = []
for date in anchor_get:
    anchor_dates.append(date.text)

# get expected departure date
depart_terminal_get = driver.find_elements(By.XPATH,"//tr[@class='color_planned']/td[9]")
depart_dates_terminal = []
for date in depart_terminal_get:
    depart_dates_terminal.append(date.text)

# create dataframe for Dubai port (based on Busan port terminal's website)
df_terminal_dubai = pd.DataFrame({'vessel': vessel_terminal, 'anchor': anchor_dates, 'departure_terminal': depart_dates_terminal})
df_terminal_dubai

# merge ONE shipper's and Busan terminal's info into one DataFrame
df_dubai = df_shipper_dubai.merge(df_terminal_dubai, on='vessel', how='left')
df_dubai


################### BELOW IS FOR DESTINATION: JEDDAH ############################
# open ONE shipper's website
driver.get("https://ecomm.one-line.com/one-ecom")
time.sleep(3)

# type in origin (input-based dropdown)
origin = driver.find_element(By.ID,"downshift-1-input")
origin.send_keys("PUSAN")
time.sleep(1)
origin.send_keys(Keys.DOWN)
origin.send_keys(Keys.RETURN)

# type in destination (input-based dropdown)
destination = driver.find_element(By.ID,"downshift-2-input")
destination.send_keys("JEDDAH")
time.sleep(1)
destination.send_keys(Keys.DOWN)
destination.send_keys(Keys.RETURN)

# next (how many weeks to search for)
driver.find_element(By.XPATH,"//div[@class='Input_inputContainer__TLu7L']/input[@data-cy='port-next-value']").click()

# select next 8 weeks
driver.find_element(By.ID,"next-week-value-56").click()

# click 'Search'
driver.find_element(By.ID,"schedule-box-search-btn-id").click()
time.sleep(2)

# get departure dates
depart_dates_get = driver.find_elements(By.XPATH,"//div[@class='ScheduleItem_date__WLAPW'][@data-cy='new-schedule-departure']")

depart_dates = []
for date in depart_dates_get:
    depart_dates.append(date.text)

# get arrival dates
arrival_dates_get = driver.find_elements(By.XPATH,"//div[@class='ScheduleItem_date__WLAPW'][@data-cy='new-schedule-arrival']")

arrival_dates = []
for date in arrival_dates_get:
    arrival_dates.append(date.text)

# get vessel's names
vessel_names_get = driver.find_elements(By.XPATH,"//a[@data-cy='new-schedule-vessel-name']")

vessel_names = []
for vessel in vessel_names_get:
    vessel_names.append(' '.join(vessel.text.split()[:-1]))

# get service lane info
service_lanes_get = driver.find_elements(By.XPATH,"//a[@data-cy='new-schedule-service-land-name']")

service_lanes = []
for lane in service_lanes_get:
    service_lanes.append(lane.text)

# create dataframe for Jeddah port (based on ONE shipper's website)
df_shipper_jeddah = pd.DataFrame({'departure': depart_dates, 'arrival': arrival_dates, 'vessel': vessel_names, 'route': service_lanes})
df_shipper_jeddah

# open Busan port terminal website
driver.get("http://www.hjnc.co.kr/esvc/vessel/berthScheduleT")


# click 'manual input'
driver.find_element(By.XPATH,"//input[@name='chkPeriod'][@value='mm']").click()

# change ending date to two months laterfor search
month = driver.find_element(By.XPATH,"//select[@id='selEndMonth']")
month.send_keys(Keys.DOWN)
month.send_keys(Keys.DOWN)

# search for 'MD3' service lane (Jeddah has two lanes: MD3 and AR1)
driver.find_element(By.ID,"searchRoute").send_keys("MD3")
driver.find_element(By.ID,"btnSearch").click()


from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

wait = WebDriverWait(driver, 10)
elem= wait.until(EC.visibility_of_element_located((By.XPATH, "//tbody/tr[@role='row']/td[5]")))

# get vessel names, anchor dates, and departure dates
md3_get_vessel = driver.find_elements(By.XPATH,"//tbody/tr[@role='row']/td[5]")
md3_get_anchor = driver.find_elements(By.XPATH,"//tbody/tr[@role='row']/td[9]")
md3_get_depart = driver.find_elements(By.XPATH,"//tbody/tr[@role='row']/td[10]")

vessel_names = []
anchor_dates = []
depart_dates = []

for vessel in md3_get_vessel:
    vessel_names.append(vessel.text)
    
for date in md3_get_anchor:
    anchor_dates.append(date.text)
    
for date in md3_get_depart:
    depart_dates.append(date.text)

# create dataframe for Jeddah port, 'MD3' lane (based on Busan port terminal's website)
df_terminal_md3 = pd.DataFrame({'vessel': vessel_names, 'anchor': anchor_dates, 'departure_terminal': depart_dates})
df_terminal_md3

# now search for 'AR1' lane 
driver.find_element(By.ID,"searchRoute").clear()
driver.find_element(By.ID,"searchRoute").send_keys("AR1")
driver.find_element(By.ID,"btnSearch").click()
time.sleep(2)

# get vessel names, anchor dates, and departure dates
ar1_get_vessel = driver.find_elements(By.XPATH,"//tbody/tr[@role='row']/td[5]")
ar1_get_anchor = driver.find_elements(By.XPATH,"//tbody/tr[@role='row']/td[9]")
ar1_get_depart = driver.find_elements(By.XPATH,"//tbody/tr[@role='row']/td[10]")

vessel_names = []
anchor_dates = []
depart_dates = []

for vessel in ar1_get_vessel:
    vessel_names.append(vessel.text)
    
for date in ar1_get_anchor:
    anchor_dates.append(date.text)
    
for date in ar1_get_depart:
    depart_dates.append(date.text)

# create dataframe for Jeddah port, 'AR1' lane (based on Busan port terminal's website)
df_terminal_ar1 = pd.DataFrame({'vessel': vessel_names, 'anchor': anchor_dates, 'departure_terminal': depart_dates})
df_terminal_ar1

# concatenate 'MD3' and 'AR1' DataFrames
df_concat = pd.concat([df_terminal_md3, df_terminal_ar1])

# merge ONE shipper's and Busan terminal's info into one DataFrame
df_jeddah = df_shipper_jeddah.merge(df_concat, on='vessel', how='left')

# change 'anchor' column into datetime
df_dubai['anchor'] = pd.to_datetime(df_dubai['anchor'])
df_jeddah['anchor'] = pd.to_datetime(df_jeddah['anchor'])

# containers can enter CY three days before the anchor date, so create new column for it
df_dubai['CY_entry'] = df_dubai['anchor'].dt.date - timedelta(days=3)
df_jeddah['CY_entry'] = df_jeddah['anchor'].dt.date - timedelta(days=3)

# convert DataFrames to excel
writer =  pd.ExcelWriter('ship_schedule.xlsx')
df_dubai.to_excel(writer, sheet_name='dubai', index=False)
df_jeddah.to_excel(writer, sheet_name='jeddah', index=False)

# set column length and alignment for better readability
for column in df_dubai:
    column_length = max(df_dubai[column].astype(str).map(len).max(), len(str(column)))
    col_idx = df_dubai.columns.get_loc(column)
    writer.sheets['dubai'].set_column(col_idx, col_idx, column_length+3)
    
for column in df_jeddah:
    column_length = max(df_jeddah[column].astype(str).map(len).max(), len(str(column)))
    col_idx = df_jeddah.columns.get_loc(column)
    writer.sheets['jeddah'].set_column(col_idx, col_idx, column_length+3)

workbook = writer.book
worksheet1 = writer.sheets["dubai"]
worksheet2 = writer.sheets["jeddah"]

fmt = workbook.add_format()
fmt.set_align('right')

worksheet1.set_column('F:F', 22, fmt)
worksheet2.set_column('F:F', 22, fmt)
    
writer.save()
writer.close()

# open excel file
os.system("start EXCEL.EXE ship_schedule.xlsx");




