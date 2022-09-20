#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2019/9/23'




import os,shutil
import time

import requests
from pyquery import PyQuery as pq

import config as cfg
import common as com
import utils
import file_struct as fs
# from utils import un_rar,removeDir,drop_ppt_filter_page,ppt_to_jpg,filter_ppt_text,filter_pic


# base_url='https://ppt.lq-sf.com'
base_url='https://www.jjppt.com'

# PPT
# menu_url='{}/ppt?page=1&order=date'.format(base_url)
#
# final_url='{}/vip/download?id=33561&type=1'.format(base_url)

# WORD
# menu_url='https://www.jjppt.com/jianli?page=1&order=download'
#
# final_url='{}/vip/download?id=259&type=0'.format(base_url)


def get_final_pages(url,type_number=1):
    content=requests.get(url,headers=cfg.get_headers(),cookies=cfg.jp_cookie).content.decode('utf-8')
    html=pq(content)
    datas=html('ul.posts>li.postsItem')

    li=[]

    for i in datas.items():
        title=i('h4.imgTitle>a').text()
        post_number=i('h4.imgTitle>a').attr('href').split('/')[-1]
        # print(title)

        url='{}/vip/download?id={}&type={}'.format(base_url,post_number,type_number)
        # print(url)
        li.append((title,url))
    return li

def save_file(url,save_path,file_name,use_cookies=True,ext='.pptx'):
    if use_cookies==False:
        content=requests.get(url,headers=cfg.get_headers()).content
    else:
        content=requests.get(url,headers=cfg.get_headers(),cookies=cfg.jp_cookie).content

    file_name=file_name+ext
    path='{}/{}'.format(save_path,file_name)


    with open(path,'wb') as f:
        f.write(content)
    print('{}:写入成功'.format(file_name))

    # print(path)
    return path


def del_empty_file(path):
    for root,dirs,files in os.walk(path):
        for f in files:
            f_path=os.path.join(root,f)
            if os.path.getsize(f_path)==0:
                os.remove(f_path)
                print('{} 是空文件 已删除'.format(f))


def get_file_number(path):
    for root,dirs,files in os.walk(path):
        return len(files)

def handle_crawl(menu_url,use_cookies=True,ext='.pptx'):

        type_num=1
        if 'jianli' in menu_url:
            type_num=0

        pages=get_final_pages(menu_url,type_num)
        for title,final_url in pages:
            downloaded_files=com.get_changed(cfg.save_path,need_ext=False)
            f_name=title
            if f_name in downloaded_files:
                continue

            # if '财务管理求职简历（含自荐信与封面）003' in downloaded_files:
            #     continue




            save_file(final_url,cfg.save_path,title,use_cookies=use_cookies,ext=ext)
            path='{}/{}'.format(cfg.save_path,title+ext)
            if ext=='.pptx':
                try:
                    utils.drop_ppt_filter_page(path,only_last=True)
                except Exception as e:
                    print(e)
            # elif ext=='.docx':
            #     utils.is_docx(path)


def handle_crawl_word(menu_url,use_cookies=True):
    handle_crawl(menu_url,use_cookies,ext='.docx')



def main(start,end):
    for i in range(start,end):

        menu_url='{}/ppt?page={}&order=download'.format(base_url,i)
        # menu_url='{}/ppt?page={}&order=date'.format(base_url,i)
        handle_crawl(menu_url,use_cookies=True)

    del_empty_file(cfg.save_path)

def handle_many(start,end,use_cookies):

    for i in range(start,end):

        menu_url='{}/ppt?page={}&order=download'.format(base_url,i)
        # menu_url='{}/ppt?page={}&order=date'.format(base_url,i)
        handle_crawl(menu_url,use_cookies=use_cookies)
    del_empty_file(cfg.save_path)

def handle_many_word(start,end,use_cookies):
    for i in range(start,end):
        menu_url='{}/jianli?page={}&order=download'.format(base_url,i)
        handle_crawl_word(menu_url,use_cookies)
    del_empty_file(cfg.save_path)

def main_many(func,start,end):
    try:
        flag=True
        need_num=(end-start)*28
        count=2

        while flag:
            if count%2==0:
                func(start,end,use_cookies=False)
            else:
                func(start,end,use_cookies=True)
                time.sleep(5)

            count+=1
            if get_file_number(cfg.save_path)>=need_num:
                flag=False
    except Exception as e:
        print(e)
        func(start,end,func)

if __name__=='__main__':

    # 5秒间隔不能超过80份PPT

    # main(10,11)

    # main_many(handle_many,1150,1151)
    # main_many(handle_many_word,50,84)
    # fs.bianli(cfg.save_path)
    # utils.handle_compress(cfg.save_path)

    # utils.handle_file()
    # 文本过滤
    # utils.filter_ppt_text(cfg.save_path,cfg.ppt_word_filter)
    # utils.filter_ppt_text(cfg.ppt_test_path,cfg.ppt_word_filter)
    # utils.filter_ppt_text(cfg.save_path,cfg.title_filter)
    # utils.filter_ppt_text(cfg.save_path,cfg.content_filter)


    # PPT转图片
    utils.ppt_to_jpg(cfg.save_path)
    utils.remove_empty_dirs(cfg.pic_path)

    utils.filter_pic()
