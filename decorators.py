#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2019/10/25'


import os
from functools import wraps
import config as cfg


# def to_many(path=cfg.save_path,file_type=('.pptx','.ppt')):
#     def outer(func):
#         @wraps(func)
#         def wrapper(*args,**kwargs):
#             for root,dirs,files in os.walk(path):
#                 for f in files:
#                     if f.endswith(file_type):
#                         res=func(f,*args,**kwargs)
#
#             return wrapper
#         return outer



