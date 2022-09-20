#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2020/1/24'

import os,shutil,json,time,random,re

import requests
from pyquery import PyQuery as pq

import config as cfg
import common as com
import utils

menu_url='http://699pic.com/ppt-0-263-new-all-0-all-all-1-0-0-0-0-0-0-all-all.html'
download_url='https://proxy-tx.699pic.com/01/66/98/a_5e144cbe82851.zip?st=PHPpOWhSp0ld930-ZC3Q_g&e=1579867403&n=%E6%91%84%E5%9B%BE%E7%BD%91_401669899.zip'

def  get_download_info(url):
    li=[]

    res=requests.get(url,headers=cfg.get_headers()).content.decode('utf-8')
    html=pq(res)
    datas=html('div.list')
    for data in datas.items():
        title=data('a.imgBorder').attr('title')
        pid=data.attr('data-id')
        li.append((title,pid))
    return li

def get_download_url(pid,page,sid=0):
    url='http://699pic.com/download/getDownloadUrl'
    data=dict(pid=pid,sid=sid,page=page)

    headers=cfg.get_headers()
    referer=dict(referer='https://699pic.com/ppt-0-263-new-all-0-all-all-{}-0-0-0-0-0-0-all-all.html'.format(page),origin='https://699pic.com')

    headers.update(cfg.st_cookie)
    headers.update(referer)
    headers.update({
        'content-type':'application/x-www-form-urlencoded',
        })

    res=requests.post(url,data,headers=headers,cookies=cfg.st_cookie)
    res=json.loads(res.content)
    # print(res)
    print(res)
    url=res.get('url',None)
    if url!=None:
        url='https:'+url
    return url

def save_file(url,file_name,path=cfg.ppt_test_path):
    res=requests.get(url,headers=cfg.get_headers()).content
    ext='.zip'
    if '.rar' in url:
        ext='.rar'
    full_path=os.path.join(path,'{}'.format(file_name+ext))

    with open(full_path,'wb') as f:
        f.write(res)
    print('{}:下载成功'.format(file_name))
    return full_path

def rename_shetu(name):
    li=['摄图网','（非企业商用）']
    for i in li:
        if i in name:
            name=name.replace(i,'')
    res=re.findall('_\d+_',name)
    if res!=[]:
        res=res[0]
        name=name.replace(res,'')

    return name

def compress_shetu(file_path=cfg.ppt_test_path,file_type='ppt',remove_source=False):
    for root,dirs,files in os.walk(file_path):
        for f in files:
            if f.endswith('.zip'):
                full_path=root+'/'+f
                utils.unzip(full_path)
                compressed_dir=full_path.replace('.zip','')
                for root2,dirs,f2s in os.walk(compressed_dir):
                    for f2 in f2s:
                        if f2.endswith(('.pptx','.ppt')):
                            f2_name,ext=os.path.splitext(f2)
                            ppt_full_path=root2+'/'+f2
                            ppt_new_name_path=compressed_dir+ext
                            print(ppt_new_name_path)
                            _,new_name=os.path.split(ppt_new_name_path)
                            new_name=rename_shetu(new_name)
                            ppt_new_path=cfg.ppt_test_path+'/'+new_name
                            os.rename(ppt_full_path,ppt_new_name_path)
                            shutil.move(ppt_new_name_path,ppt_new_path)
                            if remove_source:
                                utils.removeDir(compressed_dir)
                                os.remove(full_path)
    print('解压缩完成')


def download_file(start,end):
    for i in range(start,end):
        menu_url='https://699pic.com/ppt-0-263-new-all-0-all-all-{}-0-0-0-0-0-0-all-all.html'.format(i)
        li=get_download_info(menu_url)
        page=i
        for title,pid in li:
            # print(title,pid)
            downloaded_files=com.get_changed(cfg.ppt_test_path,need_ext=False)
            if title in downloaded_files:
                continue

            try:
                download_url=get_download_url(pid,page)
                ppt_path_zip=cfg.ppt_test_path+'/'+title+'.zip'
                ppt_path_rar=cfg.ppt_test_path+'/'+title+'.rar'
                if os.path.exists(ppt_path_zip) or os.path.exists(ppt_path_rar):
                    title=title+str(random.randint(1,1000))
                save_file(download_url,title,path=cfg.ppt_test_path)
            except Exception as e:
                print(e)


def handle_ppt(path=cfg.ppt_test_path):
    compress_shetu(file_path=path,remove_source=True)
    utils.many_drop_ppt_filter_page(path)
    utils.filter_ppt_text(path,cfg.ppt_word_filter)

if __name__=='__main__':
    # get_download_info(menu_url)
    # get_download_url(pid='401669899',page=1)
    # save_file(download_url,'1')
    # compress_shetu(remove_source=True)

    # 每天一页
    # name='摄图网_401901603_简约风个人作品集PPT模板（非企业商用）.pptx:已去出处信息'
    # rename_shetu(name)

    # 9 10

    utils.auto_mkdir()
    download_file(1,2)
    handle_ppt(path=cfg.ppt_test_path)

