# Copyright (C) 2024 Beksultan Artykbaev - All Rights Reserved

import sys
import os
import logging
from collections import Counter
from typing import List

import openpyxl
import pyexcel as p
import tqdm
from gooey import GooeyParser, Gooey

from lemmatization import lemmatize


logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

def get_execution_folder() -> str:
	if getattr(sys, 'frozen', False):
		# If the script is running as a bundled executable (e.g., PyInstaller)
		return os.path.dirname(sys.executable)
	else:
		# If the script is running as a standalone .py file
		return os.path.dirname(os.path.realpath(sys.argv[0]))

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def load_xls_sheet_values(xls_filepath, ranges: str, sheet_name=None, delimeter:str="; ", exclude_keywords:List[str]=list()) -> List[List[str]]:
    '''Reads given XLS(X) specific sheet and returns values of cells in given range.'''
    # Converting xls file to .xlsx because openpyxl doesn't support xls
    is_temp = False
    if xls_filepath.endswith("xls"):
        is_temp = True
        xls_filepath = xls_filepath
        p.save_book_as(file_name=xls_filepath,
                    dest_file_name=f"{xls_filepath}x")
        xls_filepath = f"{xls_filepath}x"

    # Loading workbook
    workbook = openpyxl.load_workbook(xls_filepath, True)
    
    # If specific sheet selected use it, active sheet otherwise
    if sheet_name != None:
        workbook.active = workbook[sheet_name]
    curr = workbook.active

    texts = list()

    # Extracting cell values
    for range_ in ranges.split("|"):
        for cell in curr[range_]:
            try:
                if len(exclude_keywords) >= 0:
                    for keyword in exclude_keywords:
                        if cell[0].value != None:
                            if keyword.lower() == cell[0].value.lower():
                                continue
                if cell[0].value != None:
                    texts.append(list())
                    texts[-1].append(cell[0].value.split(delimeter))
            except TypeError as e:
                pass

    workbook.close()
    if is_temp:
        os.remove(xls_filepath)

    return texts

def normalize(strings: List[List[str]], lemmatize_: bool=False) -> List[str]:
    '''Normalize strings for co-occurrence analysis.
    Converts strings in list in list to their lower-cased and lemmatized version'''
    normalized_words = list(list())
    i = 0
    arr_size = len(strings)
    if lemmatize_:
        logger.info("Normalizing strings. (Converting to lower-case and lemmatizing)")
        for line in strings:
            i += 1
            logger.info(f"Processed cell: {i}/{arr_size}")
            normalized_words.append(list())
            for text in line[0]:
                normalized_words[-1].append(lemmatize(text.lower()))
    else:
        logger.info("Normalizing strings. (Converting to lower-case)")
        for line in strings:
            i += 1
            logger.info(f"Processed cell: {i}/{arr_size}")
            normalized_words.append(list())
            for text in line[0]:
                normalized_words[-1].append(text.lower())
    
    logger.info("Successfully finished normalizing strings.")
    
    return normalized_words    

def get_active_sheetname(xls_filepath: str) -> List[str]:
    # Converting xls file to .xlsx because openpyxl doesn't support xls
    is_temp = False
    if xls_filepath.endswith("xls"):
        is_temp = True
        xls_filepath = xls_filepath
        p.save_book_as(file_name=xls_filepath,
                    dest_file_name=f"{xls_filepath}x")
        xls_filepath = f"{xls_filepath}x"

    # Loading workbook
    workbook = openpyxl.load_workbook(xls_filepath, True)
    sheetname = workbook.active.title
    workbook.close()
    if is_temp:
        os.remove(xls_filepath)
    return sheetname

@Gooey(program_name="D2 Research Maker Toolkit",
       image_dir=resource_path("icons"),
       default_size=(950,780),
       program_description="Simple co-occurrence analysis matrix generation tool.\nImport Data from XLSX.",
       menu=[{
        'name': 'Help',
        'items': [{
                'type': 'AboutDialog',
                'menuTitle': 'About',
                'name': 'D2 Research Maker Toolkit',
                'description': 'Research toolkit for universities all around the world.',
                'version': '0.1.0',
                'copyright': 'Copyright (C) 2024 Beksultan Artykbaev - All Rights Reserved',
                'website': 'mailto:artykbaev2003@gmail.com',
            }]
        }],
        optional_cols=4,
        required_cols=2,
        disable_progress_bar_animation=True, 

)
def main():
    parser = GooeyParser()
    parser.add_argument("filepath", metavar="Path to excel spreadsheet", type=str, widget="FileChooser", help="Choose XLS (Excel) file.",
                        gooey_options={
                            'wildcard':
                                "XLSX (Excel spreadsheet) (*.xlsx,*.xls)|*.xlsx;*.xls|"
                                "All files (*.*)|*.*",
                            'default_file': "Pick XLSX file",
                            'message': "Select XLSX (Excel spreadsheet) file"
                        }
                        )
    parser.add_argument("--sheet_name", metavar="Name of the sheet",help="Select the sheetname. (Leave empty to select the active spreadsheet.)")
    parser.add_argument("range", metavar="Range",type=str, help="Range of the cells that will be used in frequency analysis.\nExample: E1:E18|A6:A19, use '|' to select two ranges at once")
    parser.add_argument("--lemmatization", action='store_true', metavar="Lemmatization", widget="CheckBox", help="Groups together different inflected forms of the same word,\nfor example 'tree diseases' -> 'tree disease', 'asians' -> 'asian'\nTakes long time.", default=False)
    parser.add_argument("save_as", metavar="Save as...", help="Choose the output file name.",widget="FileSaver", default=os.path.join(get_execution_folder(),"output.xlsx"),
                        gooey_options={
                                'wildcard':
                                    "XLSX (Excel spreadsheet) (*.xlsx)|*.xlsx|"
                                    "All files (*.*)|*.*",
                                'message': "Create name for xlsx file",
                                'default_file': "output.xlsx"
                            })
    parser.add_argument("--delimeter", metavar="Delimeter for cell's data", default="; ", help="Select the delimeter between keys in cell value.",)
    parser.add_argument("--exclude_keywords", metavar="Exclude specific keywords", help="If you want to remove cells that contain one of specific keywords, write them down through commas.\nExample: Science, Climate change")
    args = parser.parse_args()
    logger.info("Starting algorithm.")
    logger.info(f'''The settings are:
    Excel spreadsheet path: {args.filepath}
    Sheet name: {get_active_sheetname(args.filepath)}
    Range: {args.range}
    Lemmatization: {args.lemmatization}
    Save As: {args.save_as}
    Delimeter: {repr(args.delimeter)}
    Keywords to exclude: {args.exclude_keywords}
''')
    print(normalize(load_xls_sheet_values(args.filepath, args.range), lemmatize_=True))

if __name__ == "__main__":
    main()
    #print(normalize(load_xls_sheet_values("savedrecs.xls", "T2:T10|Z2:Z10")))