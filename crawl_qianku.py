#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2019/10/6'

from gevent import monkey
monkey.patch_all()
import os,shutil,json
import time
import re
import random

import requests
from pyquery import PyQuery as pq

import config as cfg
import common as com
import utils
import crawl_qiantu as qt

word_path=cfg.save_path_word
ppt_path=cfg.save_path
ppt_test_path=cfg.ppt_test_path


base_url='https://588ku.com'

menu_url='https://588ku.com/resume/0-default-0-0-0-0-0-0-1/'

detail_url='{}/office/15987.html'.format(base_url)

final_url='https://dl.588ku.com/down/rar?callback=handleResponse&type=4&picid=92044&refererUrl=https%3A%2F%2F588ku.com%2Fresume%2F0-default-0-0-0-0-0-0-2%2F'

download_url='https://proxy-vip.588ku.com/Resume_psd/00/01/59/588ku_73a0d09d80ffd39d26d98f2ad9886bfb.zip?st=4ywKeGnsxX_FsvDyqfLGtg&e=1570417649'

def get_detail_li(url):
    li=[]
    res=requests.get(url,headers=cfg.get_headers(),cookies=cfg.qk_cookie).content.decode('utf-8')
    html=pq(res)
    datas=html('div.top-tle>a:nth-of-type(2)')
    for i in datas.items():
        title=i.text()

        num=i.attr('href').split('/')[-1].replace('.html','')

        li.append((title,num))
    return li

def get_final_url(num):
    final_url='https://dl.588ku.com/down/rar?callback=handleResponse&type=4&picid={}&refererUrl=https%3A%2F%2F588ku.com%2Fresume%2F0-default-0-0-0-0-0-0-{}%2F'
    final_url=final_url.format(num,random.randint(1,478))
    return final_url

def get_download_url(url,use_cookies=True):
    if use_cookies==True:
        res=requests.get(url,headers=cfg.get_headers(),cookies=cfg.qk_cookie).content.decode('utf-8')
    else:
        res=requests.get(url,headers=cfg.get_headers()).content.decode('utf-8')
    # print(res)

    download_url=re.findall('(http.+?)\"',res)
    if download_url:
        download_url_path,ext=download_url[0].replace('\\','\\\\').split('st=')

        download_url_path=download_url_path.replace('\\\\','')
        download_url=download_url_path+'st='+ext

        # print(download_url)


        return download_url
    return None

def save_file(url,path,file_name):
    res=requests.get(url,headers=cfg.get_headers()).content
    ext='.zip'
    if '.rar' in url:
        ext='.rar'
    full_path=os.path.join(path,'{}'.format(file_name+ext))

    with open(full_path,'wb') as f:
        f.write(res)
    print('{}:下载成功'.format(file_name))
    return full_path



def main(start,end,path=ppt_test_path,use_cookies=True):

        for i in range(start,end):

            # 合同
            # menu_url='https://588ku.com/resume/0-default-42-0-0-0-3-0-{}/'.format(i)

            # 人力资源
            # menu_url='https://588ku.com/resume/0-default-40-0-0-0-3-0-{}/'.format(i)


            # 计划书
            # menu_url='https://588ku.com/resume/0-default-45-0-0-0-3-0-{}/'.format(i)

            # ppt最新
            menu_url='https://588ku.com/ppt/0-new-0-0-0-0-0-0-{}/'.format(i)

            # ppt推荐
            # menu_url='https://588ku.com/ppt/0-default-0-0-0-0-0-0-{}/'.format(i)


            for file_name,detail_url_num in get_detail_li(menu_url):
                try:
                    downloaded_files=com.get_changed(ppt_test_path,need_ext=False)
                    if file_name in downloaded_files:
                        continue

                    final_url=get_final_url(detail_url_num)
                    # print(final_url)
                    download_url=get_download_url(final_url,use_cookies)
                    save_file(download_url,path,file_name)
                except Exception as e:
                    print(e)

        # main(start,end,path=path,use_cookies=True)


def qianku_rename(path=cfg.ppt_test_path):
    for root,dirs,files in os.walk(path):
        for f in files:
            if f.endswith(('.zip','rar')):
                f_name,ext=os.path.splitext(f)
                res=re.findall('_(.+)_',f_name)
                # print(res)
                if res!=[]:
                    res=res[0]
                    n_name=res+ext
                    print(n_name)
                    old_path=os.path.join(root,f)
                    new_path=os.path.join(root,n_name)
                    if os.path.exists(new_path):
                        new_path=root+'/'+res+str(random.randint(1,1000))+ext
                    os.rename(old_path,new_path)
                    print(n_name,':重命名成功')

if __name__=='__main__':


    # PPT 最新
    # 每天2页
    # main(1,3)

    # qianku_rename()
    utils.handle_compress(ppt_test_path,end_with=('.pptx','.ppt'))


    utils.many_drop_ppt_filter_page(path=ppt_test_path)
    utils.filter_ppt_text(ppt_test_path,cfg.ppt_word_filter)




