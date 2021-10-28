# -*- coding: utf-8 -*-
"""
Created on Sat Oct 16 23:18:04 2021
Last Edited on Sun Oct 24 23:51:32 2021
@author: Mady
This Scipt runs an bot in firefox that searches articles in the science direct website
It takes two inputs: Search Keyword and Year of publications. 
It also creates a log file that briefly mentions if any errors encountered
"""

import time
import sys
import datetime
import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import Workbook
import selenium
from selenium import webdriver

total = 0
errStop=False

def collect(i,  keyword, yr):
    global total
    global errStop
    path = "M:/RW/Data/Bot Collected Data/"+keyword+" "+yr+" - "+str(i)+".xlsx"
    web_input = "https://www.sciencedirect.com/search?qs="+keyword+"&date="+yr+"&articleTypes=FLA&show=25"
    #web=web_input
    driver = webdriver.Firefox(executable_path="M:/RW/Data/geckodriver.exe")
    #driver.get(web)
    articles = 1
    count = 1
    email=[]
    name=[]
    surname = []
    offset=25
    offset_at_start =(i*100) #to continue from next page after previous search ended
    web =web_input+"&offset=" + str(offset_at_start)
    offset = offset_at_start+25
    count_for_you = 1
    #while articles <5: #Test cases
    while articles <100:
        try:        
            driver.get(web)
            #while count <=5: #Test Cases
            while count <=26:
    
                if count ==3:
                    count =4
                while True:
                    time.sleep(0.1)
                    content = driver.page_source
                    soup = BeautifulSoup(content, "html.parser")
                    if "result-list-title-link u-font-serif text-s" in str(soup):
                        driver.find_element_by_xpath("//ol/li[" + str(count) + "]/div/div/h2/span").click()
                        print((total+count_for_you), driver.current_url, end='\r')
                        break
                
                while True:
                    time.sleep(0.1)
                    content = driver.page_source
                    soup = BeautifulSoup(content, "html.parser")
                    if 'span class="button-text">Share' in str(soup):
                        author_count = len(soup.find_all("svg", class_="icon icon-envelope"))
                        break
                    
                author=1
                while author_count>=author:
                    while True:
                        time.sleep(0.1)
                        envelope = driver.find_elements_by_css_selector("svg.icon.icon-envelope")[author - 1].click()
                        content = driver.page_source
                        soup = BeautifulSoup(content, "html.parser")
                        if ('class="icon icon-cross"'in str(soup)) and ("science/article" in driver.current_url):
                            mail = soup.find("div", class_="e-address")
                            email.append(mail.find("a").contents[0])
                            autor = soup.find("div", class_="author u-h4")
                            author +=1
                            if "text given-name" in str(autor):
                                name.append(autor.find("span", class_="text given-name").contents[0])
                            else:
                                name.append("none")
                            if "text surname" in str(autor):
                                surname.append(autor.find("span", class_="text surname").contents[0])
                            else:
                                surname.append("none")
                            break
                        elif "science/article" in driver.current_url:
                            continue
                        else:
                            author = author_count+1
                            break
                driver.get(web)
                count +=1
                count_for_you +=1
                writeFile(path, name, surname, email)
                
            articles = articles+26
            web = web_input+"&offset=" + str(offset)
            offset = offset +25
            count = 1
        except KeyboardInterrupt:
            errStop=True
            total = total + count_for_you
            writeFile(path, name, surname, email)
            return True
        except (selenium.common.exceptions.NoSuchWindowException,
            selenium.common.exceptions.InvalidSessionIdException,
            selenium.common.exceptions.WebDriverException):
            logError(str(sys.exc_info()))
            total = total + count_for_you
            del driver
            driver = webdriver.Firefox(executable_path="M:/RW/Data/geckodriver.exe")
            continue
        except:
            logError(str(sys.exc_info()))
            driver.quit()
            del driver
            driver = webdriver.Firefox(executable_path="M:/RW/Data/geckodriver.exe")
            continue 
    
    print(len(name),len(surname),len(email))
    total = total + len(email)
    driver.quit()
    del articles
    del count
    del email
    del name
    del surname
    #del offset
    time.sleep(10)
    return False

def logError(msg):
    f=open('LogFile.txt','a')
    f.write("---------------------------------------------------------------------------------------------------------\n")
    f.write("Error Occured!!!!!\n")
    errTime = str(datetime.datetime.now())
    f.write("Time of error = "+errTime+"\n")
    f.write("Error information is shown below........\n")
    f.write(msg)
    f.write("\n---------------------------------------------------------------------------------------------------------\n")
    f.close()
    return

def stopped(intr):
    global total
    f=open('LogFile.txt','a')
    f.write("---------------------------------------------------------------------------------------------------------\n")
    if intr:
        f.write("Collection Stopped by user !!!!!\n")
    else:
        f.write("Collection Stopped after search finished !!!!!\n")
    stopTime = str(datetime.datetime.now())
    f.write("Time = "+stopTime)
    f.write("\nTotal collected emails - "+str(total))
    f.write("\n---------------------------------------------------------------------------------------------------------\n")
    f.close()
    return

def writeFile(path, name, surname, email):
    book = Workbook()
    writer = pd.ExcelWriter(path, engine='openpyxl')
    writer.book = book
        
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        
    names_data = pd.DataFrame(name)
    surnames_data = pd.DataFrame(surname)
    email_data = pd.DataFrame(email)
    full_name = names_data+" "+surnames_data

    #names_data.to_excel(writer, "Main", startrow=1, startcol=1, index=False, header=False)
    full_name.to_excel(writer, "Main", startrow=1, startcol=1, index=False, header=False)
    #surnames_data.to_excel(writer, "Main", startrow=1, startcol=2, index=False, header=False)
    email_data.to_excel(writer, "Main", startrow=1, startcol=0, index=False, header=False)
    writer.save()
    
    del names_data
    del surnames_data
    del full_name
    del email_data
    return

def main():
    global errStop
    key = input("Enter search keyword: ")
    yr = input("Enter year of publication: ")
    i=0
    while i<10:
        err = collect(i, key, yr)
        if err:
            break
        else:
            i=i+1
    print("Total collected email IDs - ",total)
    stopped(errStop)
    return

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("Total collected email IDs - ",total)
        stopped(True)