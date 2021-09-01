from selenium import webdriver
import requests
import pandas as pd
from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

hindex_list_all = []
hindex_list_2016 = []
citations_2016 = []
citations_all = []
i10_all = []
i10_2016 = []
fac_info = []
research_interests = []
google_scholars = []

driver = webdriver.Chrome("/Users/srihari/Desktop/ZENSE_SUBMISSION/chromedriver")
webpage = 'https://www.cse.iitb.ac.in/people/faculty.php'
driver.get(webpage)
driver.implicitly_wait(5)

names = driver.find_elements(By.XPATH, "//div[@id = 'current']//div[@class = 'card']//div[@class = 'side']//div[@class = 'info']//div[@class = 'name']//a")
for name in names:
    fac_info.append(name.text)
    query = name.text + " iitb"
    url = 'https://scholar.google.com/citations?hl=en&view_op=search_authors&mauthors=' + query + '&btnG='
    result = requests.get(url).text
    result_page = BeautifulSoup(result, 'lxml')
    i = result_page.find('a', class_='gs_ai_pho')
    if (i == None):
        hindex_list_all.append(0)
        hindex_list_2016.append(0)
        citations_all.append(0)
        citations_2016.append(0)
        i10_all.append(0)
        i10_2016.append(0)
        google_scholars.append("None")
        continue
    query_new = i['href']
    new_url = 'https://scholar.google.com' + query_new
    google_scholars.append(new_url)
    final = requests.get(new_url).text
    final_page = BeautifulSoup(final, 'lxml')
    k = 0
    for j in final_page.findAll('td', class_='gsc_rsb_std'):
        k += 1
        if (k == 1):
            all_citations = j.text
        if (k == 2):
            x_citations = j.text
            citations_all.append(all_citations)
            citations_2016.append(x_citations)
        if (k == 3):
            all_hindex = j.text
        if (k == 4):
            x_hindex = j.text
            hindex_list_all.append(all_hindex)
            hindex_list_2016.append(x_hindex)
        if (k == 5):
            all_i10 = j.text
        if (k == 6):
            x_i10 = j.text
            i10_all.append(all_i10)
            i10_2016.append(x_i10)

research_info = driver.find_elements(By.XPATH, "//div[@id = 'current']//div[@class = 'card']//div[@class = 'side']//div[@class = 'info']//div[@class = 'office']//div[@class = 'email']//div[@class = 'body']")
for info in research_info:
    research_interests.append(info.text)

research_list = research_interests[:123:3]

driver.close()

df = pd.DataFrame()
df['Name'] = fac_info
df['All citations'] = citations_all
df['citations after 2016'] = citations_2016
df['total h index'] = hindex_list_all
df['h index after 2016'] = hindex_list_2016
df['all i10'] = i10_all
df['i10 after 2016'] = i10_2016
df['Research Interests'] = research_list
df['Google Scholar Link'] = google_scholars

df.to_excel('final_iitb.xlsx', index = False)

writer = pd.ExcelWriter('final_iitb.xlsx')
df.to_excel(writer, sheet_name='final', index=False, na_rep='NaN')

# Auto-adjust columns' width
for column in df:
    column_width = max(df[column].astype(str).map(len).max(), len(column))
    col_idx = df.columns.get_loc(column)
    writer.sheets['final'].set_column(col_idx, col_idx, column_width)

writer.save()


