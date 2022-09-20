#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2019/9/15'

import os,shutil

import requests
from pyquery import PyQuery as pq

from config import save_path,headers,ppt_word_filter,can_use_path
from utils import un_rar,removeDir,drop_ppt_filter_page,ppt_to_jpg,filter_ppt_text,filter_pic
import common as com

base_url="http://www.ypppt.com"
content_filter=['设计']

# download_url='{}/p/d.php?aid=5645'.format(base_url)
#
# final_url='http://www.youpinppt.com/soft/190805/1-1ZP5100058.rar'
#
#
# file_name='测试用PPT'

def get_download_page(url):
    content=requests.get(url,headers=headers).content.decode('utf-8')
    html=pq(content)
    datas=html('ul.posts>li')
    li=[]
    for item in datas.items():
        title=item('a.p-title').text()
        url=item('a.p-title').attr('href')
        post_number=url.split('/')[-1].replace('.html','')
        download_url='{}/p/d.php?aid={}'.format(base_url,post_number)
        li.append((title,download_url))
    return li

def get_final_page(url):
    content=requests.get(url,headers=headers).content.decode('utf-8')
    html=pq(content)
    final_url=html('div.box>ul.down>li:nth-of-type(1)>a').attr('href')

    if 'pan.baidu' in final_url:
        return None

    elif not final_url.startswith('http://'):
        final_url=base_url+final_url
    return final_url

def save_file(url,save_path,file_name):
    content=requests.get(url,headers=headers).content

    file_name=file_name+'.rar'
    path='{}/{}'.format(save_path,file_name)


    with open(path,'wb') as f:
        f.write(content)
    print('{}:写入成功'.format(file_name))

    print(path)
    return path

def filter_file(path,rename):
    try:
        for root,dirs,files in os.walk(path):
            print(dirs)
            for i in files:
                if i.endswith(('.ppt','.pptx')):
                    print(i)
                    full_path=os.path.join(path,i)
                    break
        # os.rename(file_path,rename_path)
        print(save_path)
        new_file_path=save_path+'/'+rename+'.pptx'
        shutil.move(full_path,new_file_path)
        os.remove(save_path+'/'+rename+'.rar')

        print('{}:过滤成功'.format(rename))
        removeDir(save_path+'/'+rename)
        return new_file_path
    except Exception as e:
        print(e)
    finally:
        removeDir(save_path+'/'+rename)



def main(menu_url):

    download_urls=get_download_page(menu_url)
    for i in download_urls:
        file_name,download_url=i
        final_url=get_final_page(download_url)
        downloaded_files=com.get_changed(save_path)
        f_name='{}.pptx'.format(file_name)
        if f_name in downloaded_files:
            continue
        if final_url==None:
            continue
        else:
            path=save_file(final_url,save_path,file_name)
            un_rar(path)
            dir_name=path.replace('.rar','')
            try:
                ppt_path=filter_file(dir_name,file_name)
                drop_ppt_filter_page(ppt_path)

            except Exception as e:
                print(e)
                continue

if __name__=='__main__':
    # for i in range(109,114):
    #     menu_url='{}/moban/list-{}.html'.format(base_url,i)
    #     main(menu_url)

    # filter_ppt_text(can_use_path,ppt_word_filter)
    # filter_ppt_text(save_path,ppt_word_filter)
    # filter_ppt_text(save_path,content_filter)

    # ppt_to_jpg(can_use_path)
    ppt_to_jpg(save_path)
    filter_pic()
