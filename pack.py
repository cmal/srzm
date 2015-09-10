#coding: utf-8
from distutils.core import setup
import py2exe
setup(console=["srzm.py"],
data_files=[(".",["file_list.txt"])]
)
