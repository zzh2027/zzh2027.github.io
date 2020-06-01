from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import ElementClickInterceptedException,ElementNotInteractableException, NoSuchElementException, StaleElementReferenceException
import base64
import requests
from requests.exceptions import InvalidSchema
import time
import json
import os
import glob
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
import numpy as np
import re
import os
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC 
from selenium.webdriver.common.by import By

class ScrapeGoogle(object):
    
    def __init__(self, headless = None, web_root = r"C:\Users\huyia\1_jupyter\爬数据爬图片", pic_root = r"C:\Users\huyia\OneDrive\Pictures"):
        
        os.chdir(web_root)
        self.web_root = web_root
        self.pic_root = pic_root
        from selenium.webdriver.chrome.options import Options
        chrome_options = Options()
        if headless == 1:
            chrome_options.add_argument('--headless')
        driver = webdriver.Chrome(chrome_options=chrome_options)
        self.driver = driver

        
        
    def __clickable(self, elem_id, typ = 'id'):
        
        ## wait until the button can be clicked
        driver = self.driver
        wait = WebDriverWait(driver, 10)
        actions = ActionChains(driver)
        if typ == 'id':
            element = wait.until(EC.element_to_be_clickable((By.ID, elem_id)))
        elif typ == 'path':
            element = wait.until(EC.element_to_be_clickable((By.XPATH, elem_id)))
        actions.move_to_element(element).perform()
        driver.execute_script("arguments[0].click();", element)   
    
    def __rollandclick(self, button):
        from selenium.webdriver.common.action_chains import ActionChains
        try:
            button.click()
        except Exception as e:
            ActionChains(self.driver).key_down(Keys.DOWN).perform()
            self.__rollandclick(button)
        
    def getPic(self, topic = 'Avengers', num_pic = 50, pic_size = 'large',
                     url = "https://www.google.com/?&bih=937&biw=1920&hl=en"):
        
        driver = self.driver
        driver.get(url)
        driver.maximize_window()
        driver.implicitly_wait(5)
        os.chdir(self.pic_root)
        try:
            os.makedirs(topic)
        except FileExistsError:
            pass
        os.chdir(self.pic_root + "\\" + topic)
        self.download = os.path.abspath('.')
        ## 在Google搜索框里输入内容
        enter = "q"
        button = driver.find_element_by_name(enter)
        button.send_keys(topic)
        button.send_keys(Keys.ENTER)
        ##进入图片搜索结果
        image_path ='//*[@id="hdtb-msb-vis"]/div[2]'
        driver.find_element_by_xpath(image_path).click()
        ##tools添加筛选条件
        path = """//*[@id="yDmH0d"]/div[2]/c-wiz/div[1]/div/div[1]/div[2]/div[2]/div/div"""
        driver.find_element_by_xpath(path).click()
        ## 选择size
        size_button = """//*[@id="yDmH0d"]/div[2]/c-wiz/div[2]/c-wiz[1]/div/div/div[2]/div/div[1]/div/div[1]"""
        self.__clickable(elem_id = size_button, typ = 'path')
        size = {"any":"""//*[@id="yDmH0d"]/div[2]/c-wiz/div[2]/c-wiz[1]/div/div/div[3]/div/span""", 
                "large":"""//*[@id="yDmH0d"]/div[2]/c-wiz/div[2]/c-wiz[1]/div/div/div[3]/div/a[1]""", 
                "medium":"""//*[@id="yDmH0d"]/div[2]/c-wiz/div[2]/c-wiz[1]/div/div/div[3]/div/a[2]""", 
                "icon":"""//*[@id="yDmH0d"]/div[2]/c-wiz/div[2]/c-wiz[1]/div/div/div[3]/div/a[3]"""}
        path = size[pic_size]
        driver.find_element_by_xpath(path).click()
        ##点击每个图片，然后从新的图片上获取link
        for i in range(1, num_pic+1):
            ## 图片的预览，不可下载
            try:
                prev = driver.find_element_by_xpath(f"""//*[@id="islrg"]/div[1]/div[{i}]/a[1]/div[1]/img""")
            except NoSuchElementException:
                continue
            self.__rollandclick(prev)
            time.sleep(4)##
            print('等四秒钟让图片加载完')
            ## 这是可以保存的图片链接
            path = """//*[@id="Sva75c"]/div/div/div[3]/div[2]/c-wiz/div[1]/div[1]/div/div[2]/a/img"""
            url = driver.find_element_by_xpath(path).get_attribute('src')
            def request_download(IMAGE_URL, img_name, error = False):
                """
                此方程用于通过图片链接下载到本地相应文件夹
                """
                if error:
                    r = base64.b64decode(IMAGE_URL[23:].replace("\n",""))
                    ## png无损压缩 jpg失真压缩
                    with open(f'{img_name}.png', 'wb') as f:
                        f.write(r) 
                else:
                    r = requests.get(IMAGE_URL)
                    ## png无损压缩 jpg失真压缩
                    with open(f'{img_name}.png', 'wb') as f:
                        f.write(r.content) 
            time.sleep(2)
            try:
                r = requests.get(url)
                request_download(url, i, error = False)
            except InvalidSchema:
                request_download(url, i, error = True)
        os.chdir(self.pic_root)
        driver.quit()
    def removeSmall(self):
        import matplotlib.image as mpimg
        local_folder = self.download
        ### 1.2store all image names\n",
        jpeg = glob.glob(local_folder+"\\*.jpeg")
        jpg  = glob.glob(local_folder+"\\*.jpg")
        png  = glob.glob(local_folder+"\\*.png")
        pictures = jpeg+jpg+png
        removed = 0
        for pic in pictures:
            picd = plt.imread(pic,0)
            vert, hori = picd.shape[:2]
            size = os.path.getsize(pic)/1024
            if size < 200:
                ## 小于200KB的图片就删掉
                ## Filter out small pictures less than 200 KB
                os.remove(pic)
                removed+=1

