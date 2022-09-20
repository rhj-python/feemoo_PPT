#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2019/9/11'



import time
import os
import shutil
import json

from selenium_helper import SeleniumHelper
from selenium.webdriver.support.ui import WebDriverWait,Select
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from jsonpath import jsonpath


from config import save_path,ppt_type,pic_path,YEAR,MONTH
import requests

import config as cfg
from utils import win_handle,get_file_path,time_sleep_calc


def ppt_cls_handle(file_name):
    for k,v in ppt_type.items():
        words,title=v
        # print(words,title)
        for w in words:
            if w in file_name:
                return title
    return ('工作汇报','商业')


upload_url='https://www.feimaoyun.com/publishresources'


# def send_to_feemoo(browser,path):
#     browser.get(upload_url)
#     upload_css='div.upload-top-nav>span.upload-btn'
#
#     time.sleep(1)
#
#     upload_ele=browser.find_element_by_css_selector(upload_css)
#     upload_ele.click()
#
#     ppt_css='div.content>div.sub-cld:nth-of-type(3)'
#
#     wait=WebDriverWait(browser,10)
#     wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,ppt_css)))
#     ppt_ele=browser.find_element_by_css_selector(ppt_css)
#     ppt_ele.click()
#
#
#     upload_file_css='.dt-file'
#     upload_ele=browser.find_element_by_css_selector(upload_file_css)
#     upload_ele.click()
#     time.sleep(0.5)
#     win_handle(path)
#     time.sleep(0.5)
#
#     next_step_css='div.step-box1 span.next-btn'
#     next_step_ele=browser.find_element_by_css_selector(next_step_css)
#     next_step_ele.click()
#     time.sleep(1)
#
#     _,file_name=os.path.split(path)
#     ppt_type=ppt_cls_handle(file_name)
#
#     choice_css='div.step-box2 div.el-select'
#     choice_ele=browser.find_element_by_css_selector(choice_css)
#     choice_ele.click()
#     time.sleep(0.5)
#
#     select_target_xpath='//li/span[contains(text(),"{}")]'.format(ppt_type)
#     select_file_ele=browser.find_element_by_xpath(select_target_xpath)
#     select_file_ele.click()
#
#     jpg_dir_name=file_name.replace('.pptx','')
#     jpg_dir_name=jpg_dir_name.replace('.ppt','')
#
#     jpg_paths='{}/{}'.format(pic_path,jpg_dir_name)
#     jpg_paths=get_file_path(jpg_paths,end_with=('.jpg'),to_win32=True)[:5]
#
#     for jpg_path in jpg_paths:
#         time.sleep(0.5)
#         upload_jpg_css='#softimgupload>span.el-icon-plus'
#         upload_jpg_ele=browser.find_element_by_css_selector(upload_jpg_css)
#         upload_jpg_ele.click()
#         time.sleep(0.5)
#         win_handle(jpg_path)
#     time.sleep(1)
#
#     submit_css='div.step-box2 span.next-btn'
#     submit_ele=browser.find_element_by_css_selector(submit_css)
#     submit_ele.click()
#
#     # 小于5M等待6秒
#     if os.path.getsize(path)>=200000000:
#         time.sleep(60)
#     if os.path.getsize(path)>=100000000:
#         time.sleep(40)
#     if os.path.getsize(path)>=40000000:
#         time.sleep(20)
#     if os.path.getsize(path)>=20000000:
#         time.sleep(15)
#     elif os.path.getsize(path)>=5000000:
#         time.sleep(10)
#     else:
#         time.sleep(6)
#
#     dst_path='g:/资源/PPT资源/{}/{}/{}'.format(YEAR,MONTH,file_name)
#     shutil.move(path,dst_path)
#
#     # print('{}:上传成功'.format(file_name))
#
#     browser.get('https://www.baidu.com/')
#     time.sleep(2)


def get_uploaded(start,end):
    li=[]
    for i in range(start,end):
        url='https://www.feimaoyun.com/index.php/jx/userjxfile'
        status=[1,0,-1]
        for k in status:
            data=dict(pg=i,status=k)
            res=requests.post(url,headers=cfg.fm_headers,data=data).text

            res=json.loads(res)
            print(res)
            datas=jsonpath(res,'$..data..files')[0]

            for data in datas:
                ppt_name=data['name']
                if ppt_name.endswith(('.pptx','.ppt')):
                    li.append(ppt_name)
    li=list(set(li))
    return li

def need_upload_file(start,end):
    has_uploaded_li=get_uploaded(start,end)
    for root,dirs,files in os.walk(save_path):
        for f in files:
            if f in has_uploaded_li:
                full_path=root+'/'+f
                dst_path='g:/资源/PPT资源/{}/{}/{}'.format(YEAR,MONTH,f)
                shutil.move(full_path,dst_path)
    print('过滤完成！')


def send_to_new_feemoo(browser,path):
    _,file_name=os.path.split(path)
    section_cls,main_cls=ppt_cls_handle(file_name)
    # print(section_cls,main_cls)

    browser.implicitly_wait(5)

    browser.get(upload_url)
    time.sleep(1)

    # 点击分类
    choice_1='//input[@placeholder="{}"]'.format('请选择')
    browser.find_element_by_xpath(choice_1).click()

    choice_2='//span[contains(text(),"{}")]'.format('PPT模板')
    browser.find_element_by_xpath(choice_2).click()


    choice_3='//span[contains(text(),"{}")]'.format(main_cls)
    browser.find_element_by_xpath(choice_3).click()

    choice_4='//span[contains(text(),"{}")]'.format(section_cls)
    browser.find_element_by_xpath(choice_4).click()



    title_xpath='请输入标题'
    browser.find_element_by_xpath('//input[@placeholder="{}"]'.format(title_xpath)).click()
    # time.sleep(0.5)

    upload_css='#softdocupload'
    # time.sleep(0.5)

    upload_ele=browser.find_element_by_css_selector(upload_css)
    upload_ele.click()

    time.sleep(0.5)
    win_handle(path)
    time.sleep(0.5)



    # choice_css='div.navList>div:nth-of-type({})>a'.format(main_cls)
    # choice_ele=browser.find_element_by_css_selector(choice_css)
    # ActionChains(browser).move_to_element(choice_ele).perform()
    #
    # select_target_xpath='//li[contains(text(),"{}")]'.format(section_cls)
    # select_file_ele=browser.find_elements_by_xpath(select_target_xpath)
    #
    #
    # for e in select_file_ele:
    #     try:
    #         e.send_keys(Keys.ENTER)
    #         # e.click()
    #     except Exception as e:
    #         print(e)


    time.sleep(0.5)

    js1="""
        var q=document.
        getElementById('softimgupload').setAttribute('style','width:580px!important')
        """

    browser.execute_script(js1)


    # print(file_name)

    jpg_dir_name=file_name.replace('.pptx','')
    jpg_dir_name=jpg_dir_name.replace('.ppt','')

    jpg_paths='{}/{}'.format(pic_path,jpg_dir_name)
    jpg_paths=get_file_path(jpg_paths,end_with=('.jpg','.JPG'),to_win32=True)[:5]




    for jpg_path in jpg_paths:
        time.sleep(0.5)
        upload_jpg_css='#softimgupload'
        upload_jpg_ele=browser.find_element_by_css_selector(upload_jpg_css)
        upload_jpg_ele.click()
        time.sleep(0.5)
        win_handle(jpg_path)
    time.sleep(1)

    submit_xpath='//span[contains(@class,"{}")]'.format('pushBtn')
    submit_ele=browser.find_element_by_xpath(submit_xpath)
    submit_ele.click()


    success_xpath='//span[contains(text(),"{}")]'.format('完成')

    t=time_sleep_calc(path)
    wait=WebDriverWait(browser,t)
    try:
        wait.until(EC.visibility_of_element_located((By.XPATH,success_xpath)))
        dst_path='g:/资源/PPT资源/{}/{}/{}'.format(YEAR,MONTH,file_name)
        shutil.move(path,dst_path)
    except Exception as e:
        print(e)
        time_sleep_calc(t)


    browser.get('https://www.feimaoyun.com/tgy')
    time.sleep(1)


def main():
    s1=SeleniumHelper()
    browser=s1.browser
    list_path=get_file_path(save_path,to_win32=True)
    count=len(list_path)
    for i in list_path:
        # send_to_feemoo(browser,i)
        send_to_new_feemoo(browser,i)
        count-=1
        print('剩余{}：{}'.format('ppt',count))
        # time.sleep(240)

if __name__=='__main__':
    # 检查出上传失败，以便再次上传
    # need_upload_file(1,3)

   main()


