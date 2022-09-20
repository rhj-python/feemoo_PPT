#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2019/9/8'

from gevent import monkey
monkey.patch_all()
import gevent

import requests
import os
import shutil
import time
import random

import win32com
from pyquery import PyQuery as pq

from utils import unzip,removeDir,drop_ppt_filter_page,ppt_to_jpg,un_rar
from config import save_path,headers

base_url='http://www.51pptmoban.com'

# index_5 ~ index_181
menu_page='{}/zt/qitazuopin/index_{}.html'.format(base_url,5)

detail_page='{}/zhuti/9103.html'.format(base_url)
download_page='{}/e/DownSys/DownSoft/?classid=2&id=9103&pathid=0'.format(base_url)

final_page='{}/e/DownSys/GetDown/?enews=DownSoft&classid=2&id=9103&pathid=0&pass=314052be157b65488c3b90da0a3c8faf&p=:::'.format(base_url)







def get_many_detail_page(url):
    content=requests.get(url,headers=headers).content.decode('gbk')
    # print(content)
    html=pq(content)
    data=html('div.pdiv')
    li=[]
    for i in data.items():
        res=i('a').attr('href')
        res=base_url+res
        li.append(res)
    return li

def get_download_page(url):
    content=requests.get(url,headers=headers).content.decode('gbk')
    html=pq(content)
    title=html('div.title_l>h1').text()
    print(title)
    res=html('div.ppt_xz>a').attr('href')
    start_path='{}/e/'.format(base_url)
    res=res.replace('/e/',start_path)
    return (title,res)

def get_final_page(url):
    content=requests.get(url,headers=headers).content.decode('gbk')
    html=pq(content)
    res=html('div.down>a').attr('href')
    start_path='{}/e/DownSys/'.format(base_url)
    res=res.replace('../',start_path)
    return res

def save_file(url,save_path,file_name):
    content=requests.get(url,headers=headers).content

    file_name=file_name+'.rar'
    path='{}/{}'.format(save_path,file_name)


    with open(path,'wb') as f:
        f.write(content)
    print('{}:写入成功'.format(file_name))

    return path

def filter_file(path,rename):
    try:
        for root,dirs,files in os.walk(path):
            print(dirs)
            for i in files:
                if i.endswith(('.ppt','.pptx')):
                    print(i)
        print(root)
        file_path=os.path.join(root,i).replace('\\','/')
        rename_path=os.path.join(root,rename).replace('\\','/')
        os.rename(file_path,rename_path)
        print(save_path)
        new_file_path=save_path+'/'+rename+'.pptx'

        shutil.move(rename_path,new_file_path)
        os.remove(save_path+'/'+rename+'.rar')

        print('{}:过滤成功'.format(rename))
        removeDir(save_path+'/'+rename)
        return new_file_path
    except Exception as e:
        print(e)
    finally:
        removeDir(save_path+'/'+rename)


def main(menu_page):

    detail_page_list=get_many_detail_page(menu_page)
    for detail_page in detail_page_list:
        file_name,download_page=get_download_page(detail_page)
        final_page=get_final_page(download_page)

        # 已测试完成
        # --------------
        # file_name='测试用ppt'
        path=save_file(final_page,save_path,file_name)

        un_rar(path)

        dir_name=path.replace('.rar','')
        try:
            ppt_path=filter_file(dir_name,file_name)
            drop_ppt_filter_page(ppt_path,only_last=True)

        except Exception as e:
            print(e)
            continue


if __name__=='__main__':
    # for i in range(38,42):
    #     menu_page='{}/zt/qitazuopin/index_{}.html'.format(base_url,i)
    #     main(menu_page)

    # gevent.joinall(
    # [gevent.spawn(main,'{}/zt/qitazuopin/index_{}.html'.format(base_url,i)) for i in range(7,11)]
    # )

    ppt_to_jpg(save_path)
