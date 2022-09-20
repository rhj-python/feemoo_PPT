#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2020/1/25'

import os,shutil,re,random

import config as cfg
import utils

filter_li=[' ','-']

def reword(file_name,re_text=''):
    res=re.findall(re_text,file_name)
    if res!=[]:
        res=res[0]
        file_name=file_name.replace(res,'')
    return file_name

def hand_xm_ppt(path=cfg.ppt_test_path,remove_source=False):
    for root,dirs,files in os.walk(path):
        for f in files:
            if f.endswith(('.pptx','.ppt')):
                file_name,ext=os.path.splitext(f)
                full_path=root+'/'+f
                for f_word in filter_li:
                    if f_word in file_name:
                        file_name=file_name.replace(f_word,'')

                file_name=reword(file_name,'^\d+')
                file_name=reword(file_name,'\(\d+\)|（\d+）|\(\w+\)|（\w+）')
                if not '模板' in file_name:
                    file_name=reword(file_name,'\d+$')
                    file_name=file_name+'模板'


                new_name_path=root+'/'+file_name+ext
                if os.path.exists(new_name_path):
                     new_name_path=root+'/'+file_name+str(random.randint(1,1000))+ext

                os.rename(full_path,new_name_path)
                new_path=path+'/'+file_name+ext
                if os.path.exists(new_path):
                     new_path=path+'/'+file_name+str(random.randint(1,1000))+ext

                shutil.move(new_name_path,new_path)
                print('{}:预处理完成'.format(new_path))

    if remove_source:
        for root,dirs,files in os.walk(path):
            for f in files:
                if not f.endswith(('.pptx','.ppt')):
                    full_path=root+'/'+f
                    os.remove(full_path)
            for dir in dirs:
                utils.removeDir(root+'/'+dir)
        print('非相关文件已删除！')




if __name__=='__main__':
    # hand_xm_ppt(path=cfg.ppt_test_path,remove_source=True)

    # utils.many_drop_ppt_filter_page(path=cfg.ppt_test_path,only_last=-2)
    utils.filter_ppt_text(cfg.ppt_test_path,cfg.ppt_word_filter)

