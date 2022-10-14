#%%
import pandas as pd
import random
import regex as re
from bs4 import BeautifulSoup
from time import sleep
from selenium import webdriver
from openpyxl import load_workbook

def find_elements(html, tag, attr, val):
    soup = BeautifulSoup(html, 'html.parser')
    return soup.find(tag, attrs={attr:val})

def clean_tickers(tick_col):
    return [s.replace('.', ',') if '.' in s else s for s in tick_col]

def month_to_int(date):
    months = {'January':1,
              'February':2,
              'March':3,
              'April':4,
              'May':5,
              'June':6,
              'July':7,
              'August':8,
              'September':9,
              'October': 10,
              'November':11,
              'December':12
    }
    date_split = re.findall(r'[^,\s]+', date)
    return [months[i] if not i.isnumeric() else int(i) for i in date_split]

def largest_str_len(col):
    largest = 0
    for i in col:
        if len(i) > largest:
            largest = len(i)
            val = i
    return len(val.split(','))

def back_month(dt_date, back=1):
    day = dt_date.day
    month = dt_date.month
    if month > back:
        month -= back
        return month, day, dt_date.year
    else:
        return (month - back + 12), day, dt_date.year - 1

def get_sheet_names_xlsx(filepath):
    wb = load_workbook(filepath, read_only=True, keep_links=False)
    return wb.sheetnames

path = 'C:/Users/agarc/AbbyCode/data/'
#%%
sheet_names = get_sheet_names_xlsx(path+'result_test.xlsx') # exclude jan / feb
wd = webdriver.Chrome()
with pd.ExcelWriter(path+"FINAL_result_test.xlsx") as writer:

    for name in sheet_names:
        original = pd.read_excel(path+'result_test.xlsx', sheet_name=name)
        try:
            links = []
            pub_dates = []
            c = 0
            for i, tick in enumerate(original['COMPANY (TICKER)']):
                date = pd.to_datetime(original['Trade Date'])[i]
                month_min, day, yr = back_month(date, back=3) # goes back
                yr_min, yr_max = yr, yr
                month_max = (month_min + 2)

                if 10 < month_min < 13:
                    month_max = month_min - 12 + 2 # max month is 2 months after min
                    yr_max -= 1
                elif month_min == 10:
                    yr_max -= 1

                url = f'https://www.google.com/search?q={tick}+stock&rlz=1C1CHBF_enUS1024US1025&biw=1564&bih=932&sxsrf=ALiCzsaGPneyPAo-kyllnxBBtXe-FGWorQ%3A1665448856808&source=lnt&tbs=sbd%3A1%2Ccdr%3A1%2Ccd_min%3A{month_min}%2F{day}%2F{yr}%2Ccd_max%3A{month_max}%2F{day}%2F{yr_max}&tbm=nws'
                # print(url[181:]) # check dates in url
                wd.get(url)
                sleep(random.uniform(1, 4))

                page = wd.page_source
                link_elmnt = find_elements(page, tag='a', attr='class', val='WlydOe')
                try:
                    article_link = link_elmnt.get('href')
                except (AttributeError) as error:
                    print(error, '\n will append No Results instead of link')
                    links.append('No Results')
                    pub_dates.append(None)
                else:
                    links.append(article_link) 
                    pub_dates.append(link_elmnt.find('div', attrs={'class':'OSrXXb ZE0LJd YsWzw'}).text)    
                    
        finally:
            wd.quit()

        original['Links'] = links
        original['Link Publish Date'] = pub_dates
        original.to_excel(writer, sheet_name=name, index=False)

#%%