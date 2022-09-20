#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2019/11/16'

import os,time,re,random

import requests
from pyquery import PyQuery as pq


import config as cfg
import utils
import common as com

base_url='https://ibaotu.com/'


# PPT
# 商务汇报
menu_url='https://ibaotu.com/ppt/3-113-0-0-0-1.html'


final_url='https://ibaotu.com/?m=downloadopen&a=open&id=18449611&down_type=1&&attachment_id='


def get_detail_li(url):
    li=[]
    res=requests.get(url,headers=cfg.common_headers).content.decode('utf-8')
    html=pq(res)
    datas=html('ul.sucai_list>li>a.jump-details')
    for data in datas.items():
        title=data('img').attr('alt')
        title=title.replace(' ','')
        title=title.replace('图片','')
        num_id=data('img').attr('pr-data-id')
        final_url='https://ibaotu.com/?m=download&a=open&id={}&down_type=1&&attachment_id='.format(num_id)
        li.append((title,final_url))
    return li



def save_file(url,file_name):
    res=requests.get(url,headers=cfg.common_headers,cookies=cfg.bt_cookie).content

    path=cfg.save_path+'/'+file_name+'.zip'

    with open(path,'wb') as f:
        f.write(res)
    print('{}:写入成功！'.format(file_name))


def rename_zipfile(path):
    for root,dirs,files in os.walk(path):
        for f in files:
            if f.endswith(('.zip')):
                full_path=root+'/'+f
                res=re.findall('包图网_\d+',f)
                if res!=[]:
                    new_f=f.replace(res[0],'')
                    filter_word_li=[' ','+','下载']
                    for fw in filter_word_li:
                        if fw in new_f:
                            new_f=new_f.replace(fw,'')
                    new_full_path=root+'/'+new_f
                    if os.path.exists(new_full_path):
                        new_full_path=root+'/'+str(random.randint(1,1000))+new_f
                    os.rename(full_path,new_full_path)


def handle_baotu_ppt(path=cfg.ppt_test_path):

    # rename_zipfile(cfg.save_path)

    # ppt解压
    # utils.handle_compress(cfg.save_path,('pptx','ppt'))

    # word解压缩
    # utils.handle_compress(cfg.save_path,('docx','doc'))

    # utils.filter_ppt_text(cfg.save_path,cfg.ppt_word_filter)

    # utils.ppt_to_jpg(cfg.save_path)
    # utils.filter_pic()

    rename_zipfile(path)

    # ppt解压
    utils.handle_compress(path,('pptx','ppt'))

    time.sleep(1)

    # 去除最后一页
    # utils.many_drop_ppt_filter_page(path)

    utils.filter_ppt_text(path,cfg.ppt_word_filter)


def handle_baotu_word(path=cfg.word_test_path):
    rename_zipfile(path)
    utils.handle_compress(path,('docx','doc'))


def main(num):

    # 商务汇报
    menu_url='https://ibaotu.com/ppt/3-113-0-0-0-{}.html'.format(num)

    for title,final_url in get_detail_li(menu_url):
        downloaded_files=com.get_changed(cfg.save_path,need_ext=False)
        if title in downloaded_files:
            continue
        save_file(final_url,title)



if __name__=='__main__':

    # 商务汇报
    # for i in range(1,2):
    #     main(i)

    handle_baotu_ppt()

    # handle_baotu_word()

