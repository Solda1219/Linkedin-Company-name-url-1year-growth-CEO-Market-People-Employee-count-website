from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import pandas as pd
import time
import json
import os
import sys
from selenium.webdriver.support.ui import WebDriverWait
import xlrd
from pandas import ExcelWriter
from pandas import ExcelFile


def headDriver():
    options = Options()
    options.headless = False
    options.add_argument("--window-size=1920,1200")
    try:
        driver = webdriver.Chrome(
            options=options)
        agent = driver.execute_script("return navigator.userAgent")
        driver.close()
        options.add_argument("user-agent="+agent)
        driver = webdriver.Chrome(
            options=options)
        return driver
    except:
        print("You must use same chrome version with chrome driver!")
        return 0


def headlessDriver():
    options = Options()
    options.add_argument("--headless")
    options.add_argument(f"--window-size=1920, 900")
    options.add_argument("--hide-scrollbars")
    try:
        driver = webdriver.Chrome(
            options=options)
        agent = driver.execute_script("return navigator.userAgent")
        driver.close()
        options.add_argument("user-agent="+agent)
        driver = webdriver.Chrome(
            options=options)

        return driver
    except:
        print("You must use same chrome version with chrome driver!")
        return 0


def findClickY(driver):
    element = driver.find_element_by_xpath("//a[@class='next']")
    if element:
        return element
    else:
        return False

timeout= 1

class LinkedinComScraper():
    def writeCsvheader(self, filename, columns):
        try:
            os.remove(filename)
        except:
            pass
        df = pd.DataFrame(columns=columns)
        # filename= str(datetime.datetime.now()).replace(':', '-')+'.csv'
        df.to_csv(filename, mode='x', index=False, encoding='utf-8-sig')

    def saveToCsv(self, filename, newPage, columns):
        df = pd.DataFrame(newPage, columns=columns)
        print("Now items writed in csv file!")
        df.to_csv(filename, mode='a', header=False,
                  index=False, encoding='utf-8-sig')

    def readComNameFromXlsx(self, filename):
        df = pd.read_excel(filename)
        return df['INPUT']

    def scrape(self):
        names= self.readComNameFromXlsx('Example Data.xlsx')

        columns = ['Name', 'Url', 'WebSite', 'TotalEmployee', 'OneYearGrow', 'Ceo', 'FullName', 'Marketing', 'CMO', 'Content', 'Founder', 'Marketer', 'Digital']
        filename = "LinkedinCom.csv"
        self.writeCsvheader(filename, columns)
        signinUrl = "https://www.linkedin.com/uas/login"
        search_url = "https://www.linkedin.com"

        # login
        try:
            linkedin_username = "valkengoed@hotmail.com"
            linkedin_password = "Contentoo2018!"
            driver = headlessDriver()
            driver.get(signinUrl)
            time.sleep(2)
            
            username_input = driver.find_element_by_id('username')
            username_input.send_keys(linkedin_username)

            password_input = driver.find_element_by_id('password')
            password_input.send_keys(linkedin_password)
            password_input.submit()
            time.sleep(5)
        except:
            print("Ops!!! Couldn't login. Please try again later!!")
            return

        for companyName in names:
            driver.get(search_url)
            time.sleep(timeout)
            try:
                inputElement = driver.find_element_by_xpath("//input[@class= 'search-global-typeahead__input always-show-placeholder']")
                inputElement.clear()
                time.sleep(0.1)
                inputElement.send_keys(companyName)
                inputElement.send_keys(u'\ue007')
                time.sleep(timeout)
            except:
                print('input element is not working=====================')
                continue
            try:
                schBtns= driver.find_elements_by_xpath("//button[@class= 'artdeco-pill artdeco-pill--slate artdeco-pill--2 artdeco-pill--choice ember-view search-reusables__filter-pill-button']")
                for schBtn in schBtns:
                    if schBtn.get_attribute("aria-label")=='Companies':
                        schBtn.click()
                        break
            except:
                print('There is no company for this search')
                continue
            time.sleep(timeout)
            try:
                tmpUrl= driver.find_element_by_xpath("//a[@class= 'app-aware-link']")
                url= tmpUrl.get_attribute("href")
                print("url====", url)
                driver.get(url+'insights/')
            except:
                print("There is no search result for this company search")
                continue
            
            time.sleep(3*timeout)
            try:
                for x in "all":
                    read_more_res = driver.find_elements_by_xpath("//div[@class= 'org-premium-container premium-accent-bar artdeco-card']")
                    driver.execute_script("arguments[0].scrollIntoView();", read_more_res[-1])
                    time.sleep(0.1)
            except:
                print("Couldn't scroll down. but no problem")
            soup= BeautifulSoup(driver.page_source, 'html.parser')
            newPage = []
            webSite= ''
            try:
                webSite= soup.find('a', attrs= {'class': 'ember-view org-top-card-primary-actions__action'})['href']
            except:
                pass
            totalEm= ''
            try:
                totalEm= soup.find('td', attrs= {'class': 'org-insights-module__summary-block'}).text
            except:
                pass
            yearGro= ''
            try:
                yearGro= soup.find_all('span', attrs={'class': 'org-insights-change-data org-insights-change-data--increase t-16 t-black t-bold'})[1].find('span').text
            except:
                pass
            ceoName= ''
            try:
                ceoName= soup.find('div', attrs= {'class': 'org-senior-hire-card__hire-detail-name'}).find('span').text
            except:
                pass
            ceoLink= ''
            try:
                ceoLink= soup.find('a', attrs= {'class': 'org-senior-hire-card org-senior-hire-card__link ember-view'})['href']
            except:
                pass
            new= {'Name': companyName, 'Url': url+'insights', 'WebSite': webSite, 'TotalEmployee': totalEm, 'OneYearGrow': yearGro, 'Ceo': search_url+ceoLink, 'FullName': ceoName, 'Marketing': '', 'CMO': '', 'Content': '', 'Founder': '', 'Marketer': '', 'Digital': ''}
            # finding people
            peopleKeys= ['Marketing', 'CMO', 'Content', 'Founder', 'Marketer', 'Digital']
            for peopleKey in peopleKeys:
                try:
                    marketingUrl= url+'people/?keywords='+peopleKey
                    driver.get(marketingUrl)
                    time.sleep(timeout*3)
                    peopleSoup= BeautifulSoup(driver.page_source, 'html.parser')
                    peopleNode= peopleSoup.find_all('li', attrs= {'class': 'grid grid__col--lg-8 pt5 pr4 m0'})
                    marketingProfile= ''
                    for person in peopleNode:
                        try:
                            link1= person.find('a')['href']
                            link= search_url+link1
                            marketingProfile= marketingProfile+', '+ link
                        except:
                            continue
                    new[peopleKey]= marketingProfile
                except:
                    continue
            print("new", new)
            newPage.append(new)
            self.saveToCsv(filename, newPage, columns)
        print('**************************************DONE************************************************')

if __name__ == '__main__':
    scraper = LinkedinComScraper()
    scraper.scrape()
