from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotInteractableException
from difflib import SequenceMatcher
import pandas as pd
import xlsxwriter
from datetime import datetime

driver = webdriver.Chrome("./chromedriver")    # initializing the selenium chrome web driver

username = input("enter your username here : ")
password = input("enter your password here : ")

start_date = input("enter starting data of query(DD/MM/YYYY): ")
end_date = input("enter ending date of query(DD/MM/YYYY): ")

start_date = start_date.split("/")      # converting date into a list[day, month, year]
end_date = end_date.split("/")

start_day = start_date[0]
start_day = int(start_day)
start_day = str(start_day)

start_month = int(start_date[1]) - 1      # since codetantra uses month indexing of 0 to 11, 1 is subtracted from the start month
start_month = str(start_month)
start_year = start_date[2]

end_day = end_date[0]
end_day = int(end_day)
end_day = str(end_day)
end_month = int(end_date[1]) - 1
end_month = str(end_month)
end_year = end_date[2]

now = datetime.now()
year = now.strftime("%Y")
day = now.strftime("%d")
month = now.strftime("%m")
month = int(month) - 1
month = str(month)

driver.get("https://iiitb.codetantra.com/login.jsp")   # logging in to codetantra
driver.find_element_by_id("loginId").send_keys(username)
driver.find_element_by_id("password").send_keys(password)
driver.find_element_by_id("loginBtn").click()

driver.implicitly_wait(5)

# using a try except block to detect if the login was successfull. the try block searches for an element present only on the
# page after login. If login is not successfull, it raises an exception.
try:
    driver.find_element(By.XPATH, "//div[@class = 'col my-2 my-md-3 my-lg-4']//div[@class = 'card h-100']//div[@class = 'card-footer p-0']//a[text() = 'Tests']")
    print("Login successful")
except NoSuchElementException:
    print("Login failed")
    exit()

tests = driver.find_element(By.XPATH, "//div[@class = 'col my-2 my-md-3 my-lg-4']//div[@class = 'card h-100']//div[@class = 'card-footer p-0']//a[text() = 'Tests']")
driver.get(tests.get_attribute('href'))


driver.find_element(By.ID, "startDate").click()
driver.implicitly_wait(5)
driver.find_element(By.XPATH, "//div[@class = 'xdsoft_label xdsoft_month']").click()

if(start_month == month):
    try:
        driver.find_element(By.XPATH, "//div[@class = 'xdsoft_option xdsoft_current'][@data-value = '" + start_month + "']").click()
    except ElementNotInteractableException:
        driver.find_element(By.XPATH, "//div[@class = 'xdsoft_option '][@data-value = '" + start_month + "']").click()
        pass
else:
    driver.find_element(By.XPATH, "//div[@class = 'xdsoft_option '][@data-value = '" + start_month + "']").click()
driver.implicitly_wait(5)
driver.find_element(By.XPATH, "//div[@class = 'xdsoft_label xdsoft_year']").click()

if(start_year == year):
    driver.find_element(By.XPATH, "//div[@class = 'xdsoft_option xdsoft_current'][@data-value = '" + start_year + "']").click()
else:
    driver.find_element(By.XPATH, "//div[@class = 'xdsoft_option '][@data-value = '" + start_year + "']").click()

WebDriverWait(driver = driver, timeout = 15).until(EC.presence_of_element_located((By.XPATH, "//td[@data-date = '" + start_day + "'][@data-month = '" + start_month + "']")))
driver.find_element(By.XPATH, "//td[@data-date = '" + start_day + "'][@data-month = '" + start_month + "']").click()

driver.find_element(By.ID, "endDate").click()
driver.implicitly_wait(5)

end1 = driver.find_elements(By.XPATH, "//div[@class = 'xdsoft_label xdsoft_month']")
end1[1].click()

if(end_month == month):
    end1 = driver.find_elements(By.XPATH, "//div[@class = 'xdsoft_option xdsoft_current'][@data-value = '" + end_month + "']")
    try:
        end1[0].click()
    except ElementNotInteractableException:
        try:
            end1[1].click()
        except ElementNotInteractableException:
            pass

else:
    end1 = driver.find_elements(By.XPATH, "//div[@class = 'xdsoft_option '][@data-value = '" + end_month + "']")
    try:
        end1[1].click()
    except IndexError:
        end1[0].click()

driver.implicitly_wait(5)
end1 = driver.find_elements(By.XPATH, "//div[@class = 'xdsoft_label xdsoft_year']")
end1[1].click()

if(end_year == year):
    end1 = driver.find_elements(By.XPATH, "//div[@class = 'xdsoft_option xdsoft_current'][@data-value = '" + end_year + "']")
    try:
        end1[1].click()
    except IndexError:
        end1[0].click()
else:
    end1 = driver.find_elements(By.XPATH, "//div[@class = 'xdsoft_option '][@data-value = '" + end_year + "']")
    end1[1].click()

WebDriverWait(driver=driver, timeout=15).until(EC.presence_of_element_located((By.XPATH, "//td[@data-date = '" + end_day + "'][@data-month = '" + end_month + "']")))
end2 = driver.find_elements(By.XPATH, "//td[@data-date = '" + end_day + "'][@data-month = '" + end_month + "']")
try:
    end2[0].click()
except ElementNotInteractableException:
    end2[1].click()

driver.find_element_by_class_name("searchTestsBtn.btn.btn-sm.btn-success").click()
exam_names = []
exam_details = []
result_links = []
dates = []
times = []
duration = []

for exam_name in driver.find_elements(By.XPATH, "//span[@class = 'text-default']"):
    exam_names.append(exam_name.text)
for l in driver.find_elements(By.XPATH, "//dl"):
    details = l.text.replace("\n", " ")
    details = details.replace("Start Time ", "")
    details = details.replace("(India Standard Time) Duration ", "")
    details1 = details.split()
    details2 = [(details1[0] + " " + details1[1] + " " + details1[2]), details1[3], (details1[4] + " " + details1[5])]
    exam_details.append(details2)
for result in driver.find_elements(By.XPATH, "//a[@title = 'Click to see Results']"):
    query = result.get_attribute('href')
    result_links.append(query)

for row in exam_details:
    dates.append(row[0])
    times.append(row[1])
    duration.append(row[2])

marks_list = []
test_names = []
marks_list1 = []
for result_link in result_links:
    driver.get(result_link)
    test_name = driver.find_element(By.XPATH, "//span[@class = 'cl-12 col-lg-6 p-0']").text
    test_name = test_name.lstrip(" Test Name : ")
    marks_scored = driver.find_element(By.XPATH, "//div[@class = 'card-body p-1 text-center'][@id = 'userMarks']").text
    test_names.append(test_name)
    marks_list1.append(marks_scored)

for name in exam_names:
    if(test_names == []):
        marks_list.append("Not Found")
        continue
    if(SequenceMatcher(a = name, b = test_names[0]).ratio() > 0.9):
        marks_list.append(marks_list1[0])
        test_names.pop(0)
        marks_list1.pop(0)
    else:
        marks_list.append("Not Found")

percentage = []
for mark in marks_list:
    if(mark == "Not Found"):
        percentage.append("-")
    else:
        mark1 = mark.split()
        mark1.remove("/")
        percentage.append((float(mark1[0])/float(mark1[1]))*100)

driver.quit()
df = pd.DataFrame()
df['Exam Name'] = exam_names
df['Date'] = dates
df['Time'] = times
df['Duration'] = duration
df['Marks Obtained'] = marks_list
df['Percentage'] = percentage


df.to_excel('grades.xlsx', index = False)

writer = pd.ExcelWriter('grades.xlsx')
df.to_excel(writer, sheet_name='my_analysis', index=False, na_rep='NaN')

# Auto-adjust columns' width
for column in df:
    column_width = max(df[column].astype(str).map(len).max(), len(column))
    col_idx = df.columns.get_loc(column)
    writer.sheets['my_analysis'].set_column(col_idx, col_idx, column_width)

writer.save()





