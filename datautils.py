import warnings
warnings.filterwarnings('ignore')

import os
import shutil
import xlsxwriter
from distutils.dir_util import copy_tree

import numpy as np
import pandas as pd
pd.set_option('display.max_colwidth', None)
pd.set_option('display.max_columns', None)


__author__ = "Sreejith Sreekumar"
__email__ = "ssreejith@protonmail.com"
__version__ = "0.0.1"




def check_and_create_folder(foldername):

    '''
    Check if a folder exists. If it exists, delete it.
    Create a new folder in the name passed.

    :param str foldername: The name of the folder
    '''

    if os.path.exists(foldername):
        shutil.rmtree(foldername)

    os.makedirs(foldername)


def insert_excel_table(sheet,df, writer):

    '''

    Format the dataframe as an excel table.

    :param str sheet: Name of the excel sheet
    :param DataFrame df: Dataframe which needs to be put as the excel sheet
    :param ExcelWriter writer: ExcelWriter object to handle the write function

    '''
    # format as excel table (after writing data to the worksheet)
    worksheet = writer.sheets[sheet]
    worksheet.add_table(0, 0, df.shape[0], df.shape[1]-1, {
        'columns': [{'header': col_name} for col_name in df.columns]
    })


def to_excel(path, d_, sheet_name="Sheet1"):

    '''
    Write a dataframe to an excel notebook in a specified path

    :param str path: Path to the excel notebook has to be written
    :param DataFrame d_: Dataframe to be written
    :param str sheet_name: Name of the sheet
    '''

    with pd.ExcelWriter(path) as writer:

        d_.to_excel(writer, sheet_name=sheet_name, index=False)
        insert_excel_table(sheet_name, d_, writer)


def copytree(src, dst, symlinks=False, ignore=None):
    '''
    Copy a folder (and its contents) recursively to another location

    :param str src: Path of the source folder
    :param str dst: Path of the destination folder

    '''
    for item in os.listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)
        if os.path.isdir(s):
            shutil.copytree(s, d, symlinks, ignore)
        else:
            shutil.copy2(s, d)

