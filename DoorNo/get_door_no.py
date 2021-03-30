#coding: UTF-8
from selenium import webdriver
from selenium.webdriver.support.ui import Select
import pandas as pd
import time
from tqdm import tqdm
import sys
import os
import random

class get_door_no():
    def __init__(self):
        self.dir_name = 'C:/temp'
        self.csv_name = 'opendata109road.csv'

        self.index_list = [0, 6840, 13680, 20520, 27360, 34204]
        #self.index_list = [0, 5, 10, 15, 20, 34204]
        
        
    #讀入數據
    def data_processor(self):
        csv_path = os.path.join(self.dir_name, self.csv_name)
        index_down = self.index_list[self.mechine_id]
        index_up = self.index_list[self.mechine_id + 1]
        data_df = pd.read_csv(csv_path).iloc[index_down : index_up, :]
        self.save_csv = data_df
        self.save_csv['door'] = None
        print(data_df)

    def browser_chrome(self, city, site, road_name):
        #模擬瀏覽器
        browser=webdriver.Chrome("C:/Program Files/Google/Chrome/Application/chromedriver.exe")
        url='https://www.ris.gov.tw/app/portal/3053'
        #get方式進入網站
        browser.get(url)
        #網站有loading時間
        time.sleep(2 * random.uniform(0.5, 1))
        browser.switch_to_frame("content-frame")
        searchBtn=browser.find_element_by_xpath("//button[@title='以部分街路門牌查詢']")
        searchBtn.click()
        #網站有loading時間
        time.sleep(2 * random.uniform(0.5, 1))

        searchBtn2=browser.find_element_by_xpath("//area[@alt='{}村里街路門牌資料']".format(city))
        searchBtn2.click()

        selecttown=Select(browser.find_element_by_id("areaCode"))#日期選單定位
        selecttown.select_by_visible_text(site)#選單項目定位

        road=browser.find_element_by_name("street")
        road.clear
        road.send_keys(road_name) 

        search=browser.find_element_by_id("goSearch")
        search.click()
        #網站有loading時間
        time.sleep(2 * random.uniform(0.5, 1))

        result=browser.find_element_by_xpath("//div[@dir='ltr' and @style='text-align:right' and @class='ui-paging-info']").text
        result = result.replace('條', '')
        tmp = result.split('共')
        browser.quit()
        return tmp[1].strip()
    
    def save(self):
        save_path = os.path.join(self.dir_name, self.csv_name.replace('.csv', '{}.csv'.format(self.mechine_id)))
        self.save_csv.to_csv(save_path, header = 1, index = False, encoding = 'utf-8')
    
    def get_mechine_id(self):
        txt_path = os.path.join(self.dir_name, 'mechine_ID.txt')
        with open(txt_path, 'r') as f:
            self.mechine_id = int(f.readline())
            print(self.mechine_id)

    def run(self):
        self.get_mechine_id()

        self.data_processor()
        for i in tqdm(range(self.save_csv.shape[0]), ncols=80):
            city = self.save_csv.iloc[i, 0]
            site = self.save_csv.iloc[i, 1]
            site = site.replace(city, '')
            road_name = self.save_csv.iloc[i, 2]
            try:
                door = self.browser_chrome(city, site, road_name)
            except:
                time.sleep(5)
                try:
                    door = self.browser_chrome(city, site, road_name)
                except:
                    door = 'error'

            self.save_csv.iloc[i, 3] = door
        print(self.save_csv)
        self.save()

if __name__ == '__main__':
    get_door_no().run()