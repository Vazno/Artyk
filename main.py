# Copyright (C) 2024 Beksultan Artykbaev - All Rights Reserved

import sys
import os
import logging
import argparse
from collections import Counter
from typing import List

import openpyxl
import pyexcel as p
import tqdm
from gooey import GooeyParser, Gooey

from lemmatization import lemmatize


logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def load_xls_sheet_values(xls_filepath, range_, sheet_name=None) -> List[str]:
    '''Reads given XLS(X) specific sheet and returns values of cells in given range.'''
    # Converting xls file to .xlsx because openpyxl doesn't support xls
    is_temp = False
    if xls_filepath.endswith("xls"):
        is_temp = True
        xls_filepath = xls_filepath
        p.save_book_as(file_name=xls_filepath,
                    dest_file_name=f"tmp_{xls_filepath}x")
        xls_filepath = f"tmp_{xls_filepath}x"

    # Loading workbook
    workbook = openpyxl.load_workbook(xls_filepath, True)
    
    # If specific sheet selected use it, active sheet otherwise
    if sheet_name != None:
        workbook.active = workbook[sheet_name]
    curr = workbook.active

    # Extracting cell values
    texts = ""
    for cell in curr[range_]:
        try:
            texts += cell[0].value
        except TypeError:
            pass

    workbook.close()
    if is_temp:
        os.remove(xls_filepath)

    # Converting texts to list
    texts = texts.split("; ")
    return texts

def normalize(strings: List[str], lemmatize_: bool=False) -> List[str]:
    '''Normalize strings for co-occurrence analysis.
    Converts strings in list to their lower-cased and lemmatized version'''
    normalized_words = list()
    
    if lemmatize_:
        logger.info("Normalizing strings. (Converting to lower-case and lemmatizing)")
        for text in tqdm.tqdm(strings):
            normalized_words.append(lemmatize(text.lower()))
    else:
        logger.info("Normalizing strings. (Converting to lower-case)")
        for text in tqdm.tqdm(strings):
            normalized_words.append(text.lower())
    logger.info("\n\n\nSuccessfully finished normalizing strings.")
    return normalized_words    

def get_all_sheet_names(xls_filepath):
    # Converting xls file to .xlsx because openpyxl doesn't support xls
    is_temp = False
    if xls_filepath.endswith("xls"):
        is_temp = True
        xls_filepath = xls_filepath
        p.save_book_as(file_name=xls_filepath,
                    dest_file_name=f"tmp_{xls_filepath}x")
        xls_filepath = f"tmp_{xls_filepath}x"
    
    # Loading workbook
    workbook = openpyxl.load_workbook(xls_filepath, True)
    sheetnames = workbook.sheetnames
    workbook.close()
    return sheetnames


@Gooey(program_name="D2 Research Tool", image_dir=resource_path("icons"))
def main():
    parser = GooeyParser(description="Simple co-occurrence analysis matrix generation tool.\nImport Data from XLSX.")
    parser.add_argument("filepath", metavar="Path to excel spreadsheet", type=str, widget="FileChooser", help="Choose XLS (Excel) file.")
    parser.add_argument("sheet_name", metavar="Name of the sheet",type=str, help="(Leave empty if you want to select the active sheet)", default=None)
    parser.add_argument("range", metavar="Range",type=str, help="Range of the cells that will be used in frequency analysis.\nExample | E1:E18 or A6:A19")
    
    args = parser.parse_args()

if __name__ == "__main__":
    main()