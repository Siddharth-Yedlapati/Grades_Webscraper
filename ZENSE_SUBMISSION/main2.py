import requests
import pandas as pd
from bs4 import BeautifulSoup


profile = []
name_list = []
edu_list = []
hindex_list_all = []
hindex_list_2016 = []
citations_2016 = []
citations_all = []
i10_all = []
i10_2016 = []
research_interests = []

quote_page = ['https://www.iiitb.ac.in/faculty', 'https://www.iiitb.ac.in/faculty/2', 'https://www.iiitb.ac.in/faculty/3']

for page in quote_page:
    r = requests.get(page)
    soup = BeautifulSoup(r.content, 'lxml')

    for article in soup.findAll('div', class_='faculty-info'):
        fac_info = article.h3.a.text
        edu_info_raw = article.find('span', class_='eduction')
        fac_info2 = fac_info + " iiitb"
        query = fac_info2
        url = 'https://scholar.google.com/citations?hl=en&view_op=search_authors&mauthors=' + query + '&btnG='
        result = requests.get(url).text
        result_page = BeautifulSoup(result, 'lxml')
        i = result_page.find('a', class_='gs_ai_pho')

        if(i == None):
            name_list.append(fac_info)
            if edu_info_raw is not None:
                edu_list.append(edu_info_raw.text.strip())
            else:
                edu_list.append("Not Found")
            hindex_list_all.append('0')
            hindex_list_2016.append('0')
            citations_all.append('0')
            citations_2016.append('0')
            i10_all.append('0')
            i10_2016.append('0')
            continue

        query_new = i['href']
        new_url = 'https://scholar.google.com' + query_new
        final = requests.get(new_url).text
        final_page = BeautifulSoup(final, 'lxml')

        name_list.append(fac_info)

        if edu_info_raw is not None:
            edu_list.append(edu_info_raw.text.strip())

        else:
            edu_list.append("Not Found")

        k = 0
        for j in final_page.findAll('td', class_='gsc_rsb_std'):
            k+=1
            if(k == 1):
                all_citations = j.text
            if(k == 2):
                x_citations = j.text
                citations_all.append(all_citations)
                citations_2016.append(x_citations)
            if(k == 3):
                all_hindex = j.text
            if(k == 4):
                x_hindex = j.text
                hindex_list_all.append(all_hindex)
                hindex_list_2016.append(x_hindex)
            if(k == 5):
                all_i10 = j.text
            if(k == 6):
                x_i10 = j.text
                i10_all.append(all_i10)
                i10_2016.append(x_i10)

df = pd.DataFrame()
df['Name'] = name_list
df['Education Info'] = edu_list
df['All citations'] = citations_all
df['citations after 2016'] = citations_2016
df['total h index'] = hindex_list_all
df['h index after 2016'] = hindex_list_2016
df['all i10'] = i10_all
df['i10 after 2016'] = i10_2016

df.to_excel('final.xlsx', index = False)

writer = pd.ExcelWriter('final.xlsx')
df.to_excel(writer, sheet_name='final', index=False, na_rep='NaN')

# Auto-adjust columns' width
for column in df:
    column_width = max(df[column].astype(str).map(len).max(), len(column))
    col_idx = df.columns.get_loc(column)
    writer.sheets['final'].set_column(col_idx, col_idx, column_width)

writer.save()





