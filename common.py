#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2019/9/26'

import os

import re

import config as cfg



def get_changed(path,choice='files',need_ext=True):
    li=[]
    for root,dirs,files in os.walk(path):
        if choice=='files':
            if need_ext==False:
                for f in files:
                    file_name,ext=os.path.splitext(f)
                    li.append(file_name)
            else:
                li.extend(files)
        elif choice=='dirs':
            li.extend(dirs)

    return li



def choice_text_to_replace(text):
    word=''
    len_text=len(text)


    # len_large_text=len(cfg.replace_large)
    # # print(len_text)
    # if len_text > len_large_text:
    # res=re.findall('\w+',text)
    # if res:
    #     len_text=len_text//3
    #
    #     count=len_text//len_large_text
    #     num=len_text%len_large_text
    #     word=cfg.replace_large*count+cfg.replace_large[:num]
    #
    # elif len_text>=20 and len_text <= len_large_text:
    #     word=cfg.replace_large[:len_text]
    # elif len_text >= 7 and len_text < 20:
    #     word=cfg.replace_medium
    # elif len_text < 7:
    #     word=cfg.replace_small
    word=cfg.replace_small*len_text

    return word


if __name__=='__main__':
    text='Liquorice, chupa chups applicake apple pie cupcake brownie bear claw gingerbread cotton candy. Bear claw croissant apple pie. Croissant cake tart liquorice tart pastry. iscuit wafer sweet apple pie. Bear claw dragée sweet roll gummi fruitcake soufflé sweet gummies pie fruitcake cotton candy sugar plum chocolate cake bears sweet roll sesame snaps topping candy sugar plum chocolate bar fruitcake soufflé sweet gummies fruitcake cotton candy sugar plum chocolate cake'
    # choice_text_to_replace(text)
    res=choice_text_to_replace(text)
    print(len(res))
