#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2019/10/25'


import time
import os,shutil
import random
import re

from selenium_helper import SeleniumHelper
from selenium.webdriver.support.ui import WebDriverWait,Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys


from config import save_path,ppt_type,pic_path,MONTH
from selenium_helper import SeleniumHelper
import config as cfg
import utils

base_url='https://www.58pic.com'

menu_url='{}/piccate/8-0-0-se1-p2.html'.format(base_url)

url='https://www.58pic.com/piccate/8-0-0-se1-p23.html'





def download_file(browser,url):

    for i in range(16,37):

        browser.get(url)
        img_box_css='.image-box'
        wait=WebDriverWait(browser,10)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,img_box_css)))
        card_css='div[key="{}"]'.format(i)
        card_ele=browser.find_element_by_css_selector(card_css)
        ActionChains(browser).move_to_element(card_ele).perform()
        time.sleep(0.5)
        button_css='{} button.download-page'.format(card_css)
        button_ele=browser.find_element_by_css_selector(button_css)
        button_ele.click()
        time.sleep(5)
        ActionChains(browser).key_down(Keys.CONTROL).send_keys("w").key_up(Keys.CONTROL).perform()
        time.sleep(1)



def main(start,end):
    browser=SeleniumHelper().browser
    for i in range(start,end):
        # 小清新
        # https://www.58pic.com/tupian/xiaoqingxin-808-0-id-0-0-%E5%B0%8F%E6%B8%85%E6%96%B0-0_2_0_0_0_0_0-0-150.html

        # 普通 23开始
        menu_url='{}/piccate/8-0-0-se1-p{}.html'.format(base_url,i)
        download_file(browser,menu_url)

def compress_zip(path,remove_source=False,rename=True,end_with=('.pptx','.ppt'),is_resume=True):
    for root,dirs,files in os.walk(path):
        for f in files:
            if f.endswith('.zip'):
                full_path=root+'/'+f
                need_name=f.replace('.zip','')
                if rename:
                    need_name=f.split('_')[1].replace('.zip','')
                # print('need_name:%s' %need_name)
                utils.unzip(full_path)
                dir_path=full_path.replace('.zip','')
                rename_file(path,dir_path,need_name,end_with,is_resume)
                if remove_source==True:
                    os.remove(full_path)

def rename_qk_ppt(path):
    for root,dirs,files in os.walk(path):
        for f in files:
            if f.endswith(('.pptx','.ppt')) and '_' in f:
                full_path=root+'/'+f
                need_name=f.split('_')[1]
                after_path=root+'/'+need_name+'.pptx'
                if os.path.exists(after_path):
                    after_path=root+'/'+need_name+str(random.randint(1,1000))+'.pptx'
                try:
                    os.rename(full_path,after_path)
                except Exception as e:
                    print(e)
                    continue

def rename_file(path,dir_name,need_name,end_with=('.pptx','.ppt'),is_resume=True):
    for root,dirs,files in os.walk(dir_name):
        for f in files:
            if f.endswith(end_with):
                try:
                    full_path=root+'/'+f
                    filter_word_li=['空白','下载','免费','的自我介绍','简历头像']

                    for fw in filter_word_li:
                        if fw in need_name:
                            need_name=need_name.replace(fw,'')

                    if is_resume==True:
                        if not '简历' in need_name:
                            need_name=need_name+'个人简历'
                    new_path=root+'/'+need_name+end_with[0]


                    os.rename(full_path,new_path)
                    move_after_path=path+'/'+need_name+end_with[0]
                    if os.path.exists(move_after_path):
                        move_after_path=path+'/'+need_name+str(random.randint(1,1000))+end_with[0]
                    shutil.move(new_path,move_after_path)
                    utils.removeDir(dir_name)
                except Exception as e:
                    print(e)


def rename_ppt(path):
     for root,dirs,files in os.walk(path):
        for f in files:
            try:
                full_path=root+'/'+f
                if f.endswith(('pptx','ppt')):
                    res=re.findall('^\d+',f)
                    if res!=[]:
                        res=res[0]
                        new_name=f.replace(res,'')
                        new_path=root+'/'+new_name
                        if os.path.exists(new_path):
                            new_file_name,ext=os.path.splitext(new_name)
                            new_name=new_file_name+str(random.randint(1,1000))+ext
                            new_path=root+'/'+new_name

                        os.rename(full_path,new_path)
                        print('{}:改名成功！'.format(f))
            except Exception as e:
                print(e)

if __name__=='__main__':

    # rename_ppt(cfg.ppt_test_path)

    # rename_qk_ppt(cfg.ppt_test_path)

    # ppt
    compress_zip(cfg.ppt_test_path,remove_source=True,is_resume=False)
    #
    # word
    # compress_zip(cfg.word_test_path,remove_source=True,end_with=('.docx','.doc','.wps'))

    # utils.many_drop_ppt_filter_page(cfg.ppt_test_path)
    # utils.filter_ppt_text(cfg.ppt_test_path,cfg.ppt_word_filter)
