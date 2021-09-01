from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import schedule
import time

username = "john.doe@iiitb.org"      # enter your codetantra username and password here
password = "john123"

def job(username, password):


    driver = webdriver.Chrome("./chromedriver")      # install chrome driver for selenium and place it in the same folder as the .py file
    driver.get("https://iiitb.codetantra.com/login.jsp")

    driver.find_element_by_id("loginId").send_keys(username)
    driver.find_element_by_id("password").send_keys(password)

    driver.find_element_by_id("loginBtn").click()       # logging in to codetantra

    driver.implicitly_wait(5)

    # using a try except block to detect if the login was successfull. the try block searches for an element present only on the
    # page after login. If login is not successfull, it raises an exception.
    try:
        driver.find_element(By.XPATH, "//div[@class = 'col my-2 my-md-3 my-lg-4']//div[@class = 'card h-100']//div[@class = 'card-footer p-0']//a[text() = 'View Classes/Meetings']")
        print("Login successful")
    except NoSuchElementException:
        print("Login failed")
        exit()

    classes = driver.find_element(By.XPATH, "//div[@class = 'col my-2 my-md-3 my-lg-4']//div[@class = 'card h-100']//div[@class = 'card-footer p-0']//a[text() = 'View Classes/Meetings']")
    driver.get(classes.get_attribute('href'))

    list = []
    lnks = driver.find_elements_by_class_name(('fc-time-grid-event.fc-event.fc-start.fc-end'))    # finding the links to all the classes in one day
    for lnk in lnks:
        list.append(lnk.get_attribute('href'))

    for link in list:
        driver.get(link)
        # using a try except block to look for the join button on each of the links. If the meeting has been started, it will
        # join the meeting through the link from the join button. If the meeting has not yet been started, it will pass and try
        # the next link.
        try:
            y = driver.find_element(By.XPATH, "//a[@class = 'btn btn-primary btn-block btn-sm']")
            driver.get(y.get_attribute('href'))
            break
        except NoSuchElementException:
            pass

# using schedule library to run the script at pre determined times every day. The times have been coded as the
# starting times for the classes of the imt2020 batch, they can be coded differently for each batch

schedule.every().day.at("09:30").do(job, username, password)
schedule.every().day.at("09:35").do(job, username, password)
schedule.every().day.at("11:30").do(job, username, password)
schedule.every().day.at("11:35").do(job, username, password)
schedule.every().day.at("13:30").do(job, username, password)
schedule.every().day.at("13:35").do(job, username, password)
schedule.every().day.at("16:00").do(job, username, password)
schedule.every().day.at("16:05").do(job, username, password)

while True:
    schedule.run_pending()
    time.sleep(1)





