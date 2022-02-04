
# -*- coding: utf-8 -*-
"""
Created on Wed Dec  2 09:39:53 2020

@author: medialytics ninja
"""
# In[ ]:
import pandas as pd
import re

#import all phone numbers
print('Importing phone numbers')
contacts = pd.ExcelFile(r'contacts.xlsx')
contacts = contacts.parse(0)
contacts['Phone'] = contacts['Phone'].astype(str)
country_code = contacts['country_code'][0].astype(int).astype(str)

#import all values from the keys.xlsx sheet
keys = pd.ExcelFile(r'keys.xlsx')
keys = keys.parse(0)

message = list(keys.type_messages_below)
message = [x for x in message if str(x) != 'nan']

image_filepath = list(keys.image_filepath)
image_filepath = [x for x in image_filepath if str(x) != 'nan']

video_filepath = list(keys.video_filepath)
video_filepath = [x for x in video_filepath if str(x) != 'nan']

document_filepath = list(keys.document_filepath)
document_filepath = [x for x in document_filepath if str(x) != 'nan']

print('Imported', contacts.shape[0],'phone numbers')

# In[]:
#Remove whitespace
print('Cleaning phone numbers')
def rem_space(text):
    text = re.sub(r'[\s]+', '', text)
    return text

contacts['Phone'] = contacts['Phone'].apply(rem_space)

# In[]:
#Extract 10 digits
contacts['Phone'] = contacts['Phone'].str[-10:]
x= contacts.shape[0]

#delete duplicates
print('Removing duplicates')
contacts.drop_duplicates(keep='first', inplace=True)
y= contacts.shape[0]
z=x-y
print(z,'duplicates found & removed')


#Add 91
contacts['Phone']= country_code + contacts['Phone'].astype(str)

print('Sending Whatsapp message to',y,'phone numbers')
# In[]:
#Just get phone nos.
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import time

df = pd.DataFrame(columns=['Phone','Status'])

print('Opening WhatsApp web')
driver = webdriver.Chrome('chromedriver.exe')
driver.get('https://web.whatsapp.com')

a=0
phone = contacts['Phone']
for i in phone:
    print('-----------------------------------------------')
    print('Sending message to: +',i)
    link = 'https://web.whatsapp.com/send?phone='+i
    driver.get(link)
    a=a+1
    df.at[a, 'Phone'] = i
    try:
        for j in message:
            print('Typing message:',j)
            inp_xpath = '//div[@title = "Type a message"]'
            input_box = WebDriverWait(driver,40).until(lambda driver: driver.find_element_by_xpath(inp_xpath))
            time.sleep(2)
            input_box.send_keys(j + Keys.ENTER)
            time.sleep(2)
            print('Message sent')
            df.at[a, 'Status'] ='Message sent succesfully'
        if image_filepath:
            for k in image_filepath:
                print('Sending image')
                attach_button_xpath = '//div[@title = "Attach"]'
                attach_button = WebDriverWait(driver,20).until(lambda driver: driver.find_element_by_xpath(attach_button_xpath))
                time.sleep(2)
                attach_button.click()
                image_box_xpath = '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]'
                image_box = WebDriverWait(driver,20).until(lambda driver: driver.find_element_by_xpath(image_box_xpath))
                image_box.send_keys(k)
                time.sleep(3)
                sendbutton = driver.find_element_by_class_name("_1w1m1")
                sendbutton.click()
                time.sleep(5)
                print('Image Sent')
                df.at[a, 'Status'] ='Message sent succesfully'
        if video_filepath:
            for k in video_filepath:
                print('Sending video')
                attach_button_xpath = '//div[@title = "Attach"]'
                attach_button = WebDriverWait(driver,20).until(lambda driver: driver.find_element_by_xpath(attach_button_xpath))
                time.sleep(2)
                attach_button.click()
                image_box_xpath = '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]'
                image_box = WebDriverWait(driver,20).until(lambda driver: driver.find_element_by_xpath(image_box_xpath))
                image_box.send_keys(k)
                time.sleep(3)
                sendbutton = driver.find_element_by_class_name("_1w1m1")
                sendbutton.click()
                time.sleep(20)
                print('Video sent')
                df.at[a, 'Status'] ='Message sent succesfully'
        if document_filepath:
            for k in document_filepath:
                print('Sending document')
                attach_button_xpath = '//div[@title = "Attach"]'
                attach_button = WebDriverWait(driver,20).until(lambda driver: driver.find_element_by_xpath(attach_button_xpath))
                time.sleep(2)
                attach_button.click()
                image_box_xpath = '//input[@accept="*"]'
                image_box = WebDriverWait(driver,20).until(lambda driver: driver.find_element_by_xpath(image_box_xpath))
                image_box.send_keys(k)
                time.sleep(3)
                sendbutton = driver.find_element_by_class_name("_1w1m1")
                sendbutton.click()
                time.sleep(20)
                print('Document sent')                
                df.at[a, 'Status'] ='Message sent succesfully'               
    except:
        print('This number is not on WhatsApp: +',i)
        df.at[a, 'Status'] ='Message Not Send'
        print('-----------------------------------------------')
            
        
print('-----------------------------------------------') 
print('Whatsapp messages sent. Exporting report') 
from datetime import datetime
dateTimeObj = datetime.now()
timestampStr = dateTimeObj.strftime("%Y%m%d_%H%M")
filename = timestampStr + '_report'
df.to_excel('{}.xlsx'.format(filename),index=False)
print('Report file exported') 
print('Program closing')    
driver.quit()
