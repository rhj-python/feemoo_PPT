#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2020/8/2'

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

from config import save_path,ppt_type_overainbow,pic_path,YEAR,MONTH
import requests
import re
from jsonpath import jsonpath
from lxml import html

import config as cfg
from utils import win_handle,get_file_path


detail_url='https://www.feimaoyun.com/#/jx/yyk18ief'

def ppt_cls_handle(file_name):
    for k,v in ppt_type_overainbow.items():
        words,title=v
        # print(words,title)
        for w in words:
            if w in file_name:
                return title
    return ('工作汇报',7)

def get_uploaded_list(page):
    li=[]

    url='https://www.feimaoyun.com/index.php/statistics/resourcelist'

    data=dict(pg=page,status='1',type=1)
    res=requests.post(url,headers=cfg.fm_headers,data=data).text
    res=json.loads(res)

    # print(res)
    # base_url='https://www.feimaoyun.com/#/jx/'
    datas=res['data']['list']


    for data in datas:

        fshort=data['fshort']
        code=fshort
        ppt_name=data['name']
        cname=data['subtype_name']
        if ppt_name.endswith(('.pptx','.ppt')):
            ppt_name=ppt_name.replace('.pptx','')
            ppt_name=ppt_name.replace('.ppt','')
            content=get_detail_page(code)
            li.append((ppt_name,cname,content))
    li=list(set(li))
    # print(li)
    return li


def get_detail_page(code):
    url='https://www.feimaoyun.com/index.php/jx/detail'
    data=dict(code=code)
    r=requests.post(url,data=data,headers=cfg.fm_headers).content
    data=json.loads(r)
    # print(data)
    imgs=jsonpath(data,'$..imgs')[0]
    imgs=['[img]{}[/img]'.format(i) for i in imgs]
    img_str='\n\n'.join(imgs)


    base_url='https://www.feimaoyun.com/#/jx/'
    detail_url=base_url+code
    detail_url='{0}[color=Blue]飞猫盘下载地址（进入下载页面后点击文件下载）[/color]：\n[url={1}]{1}[/url]'.format('\n'*5,detail_url)
    content=img_str+detail_url

    return content

def get_ob_uploaded():
    li=[]
    for i in range(1,11):
        url='http://www.over-rainbow.name/forum.php?mod=forumdisplay&fid=44&page={}'.format(i)
        r=requests.get(url,headers=cfg.headers).text
        h=html.etree.HTML(r)
        # print(h)
        titles=h.xpath('//h3[@class="xw0 line-limit-length"]/a/text()')
        li.extend(titles)
    li=list(set(li))
    return li

def send_ppt(browser,subject,ppt_type,content):

    add_ppt_page='http://www.over-rainbow.name/forum.php?mod=post&action=newthread&fid=44'
    browser.get(add_ppt_page)
    browser.implicitly_wait(5)
    wait=WebDriverWait(browser,10)
    title_css='#subject'
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,title_css)))

    only_text=browser.find_element_by_css_selector('#e_switcher')
    only_text.click()

    sel_css='#typeid_ctrl'
    sel_ele=browser.find_element_by_css_selector(sel_css)
    sel_ele.click()
    time.sleep(0.5)

    code_name=ppt_cls_handle(ppt_type)[0]
    print(code_name)
    code_type=ppt_cls_handle(ppt_type)[1]+2


    option_css='#typeid_ctrl_menu li:nth-of-type({})'.format(code_type)

    option_ele=browser.find_element_by_css_selector(option_css)
    time.sleep(0.5)
    option_ele.click()
    # time.sleep(0.5)

    title=browser.find_element_by_css_selector('#subject')
    title.click()
    time.sleep(0.5)
    title.send_keys(subject)
    # time.sleep(0.5)

    # browser.switch_to.frame("e_iframe")
    body_content=browser.find_element_by_css_selector("#e_textarea")
    body_content.click()
    body_content.send_keys(content)

    # browser.switch_to.default_content()

    addition=browser.find_element_by_css_selector('#extra_additional_b')
    addition.click()

    post_button=browser.find_element_by_css_selector('button[name="topicsubmit"]')
    post_button.click()
    print('{}:发表成功！'.format(subject))
    time.sleep(1)

def upload_to_ob(browser,page):
    data = get_uploaded_list(page)

    uploaded_li=get_ob_uploaded()
    for ppt_name,cname,content in data:
        if ppt_name not in uploaded_li:

            ppt_type=cname

            subject=ppt_name
            try:
                send_ppt(browser=browser,subject=subject,ppt_type=ppt_type,content=content)
            except Exception as e:
                print(e)
                send_ppt(browser=browser,subject=subject,ppt_type=ppt_type,content=content)


def main(start,end):
    s1=SeleniumHelper()
    browser=s1.browser

    for i in range(end,start,-1):
        print(i)
        upload_to_ob(browser,i)


if __name__=="__main__":
    # download_and_save(600,700)

    # get_uploaded_list(1,3)
    # get_detail_page(code='um80icp3')

    # print('信息化财务会计论文模板' not in get_ob_uploaded())

    main(0,1)


