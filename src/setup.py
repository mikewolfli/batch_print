# -*- coding: utf-8 -*-
'''
Created on 2015年8月25日

@author: 10256603
'''
from distutils.core import setup
import encodings
import py2exe

setup(windows=[  {
                        "script":"batch_print.py",
                }])