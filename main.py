# Copyright (C) 2024 Beksultan Artykbaev - All Rights Reserved

import sys
import os
import logging
from typing import List

import openpyxl
import spacy
from gooey import GooeyParser, Gooey

# ---------------------------------------
# NOTE: Do not remove these pyexcel imports as it requires direct imports to successfully compile using pyinstaller
import pyexcel as p
import pyexcel_xls 
import pyexcel_xlsx
import pyexcel_xlsxw
import pyexcel_io
import pyexcel_io.writers
# ---------------------------------------

from matrix_generator import generate_excel, generate_co_occurrence_matrix

# Setting up logger
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

def load_xls_sheet_values(xls_filepath, ranges: str, sheet_name=None, delimeter:str="; ") -> List[List[str]]:
    '''Reads given XLS(X) specific sheet and returns values of cells in given range.'''
    # Converting xls file to .xlsx because openpyxl doesn't support xls
    is_temp = False
    if xls_filepath.endswith(".xls"):
        is_temp = True
        p.save_book_as(file_name=xls_filepath,
                    dest_file_name=f"{xls_filepath}x")
        xls_filepath = f"{xls_filepath}x"

    if xls_filepath.endswith(".csv"):
        is_temp = True
        p.save_book_as(file_name=xls_filepath,
                    dest_file_name=xls_filepath.replace(".csv", ".xlsx"))
        xls_filepath = xls_filepath.replace(".csv", ".xlsx")
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
                if cell[0].value != None:
                    texts.append(list())
                    for word in cell[0].value.split(delimeter):
                        texts[-1].append(word.strip())

            except TypeError as e:
                pass

    workbook.close()
    if is_temp:
        os.remove(xls_filepath)

    return texts

def homogenize(graph: List[List[str]], lemmatize_: bool=False, language:str="english") -> List[List[str]]:
    '''Homogenize strings for co-occurrence analysis.
    Converts strings in list in list to their lower-cased and lemmatized version'''
    homogenized_words = list(list())
    i = 0
    arr_size = len(graph)

    if language == "english":
        english = resource_path(os.path.join("models", "en_core_web_sm"))
        nlp = spacy.load(english)

    if lemmatize_:
        logger.info("(Converting data to lower-cased and lemmatizized version)")
        for line in graph:
            i += 1
            logger.info(f"Processed cell: {i}/{arr_size}")
            homogenized_words.append(list())
            for text in line:
                text = text.lower()
                doc = nlp(text)
                lemmas = [token.lemma_ for token in doc]
                s = " ".join(lemmas)
                homogenized_words[-1].append(s)
    else:
        logger.info("(Converting data to lower-cased version)")
        for line in graph:
            i += 1
            logger.info(f"Processed cell: {i}/{arr_size}")
            homogenized_words.append(list())
            for text in line[0]:
                homogenized_words[-1].append(text.lower())
    return homogenized_words

def exclude_keywords_from_graph(graph: List[List[str]], exclude_keywords: List[str]) -> List[List[str]]:
    '''Returns graph where given keywords are excluded from graph (Nodes connected to the excluded keywords (nodes) are removed too).'''
    fixed_graph = list()
    if exclude_keywords == None:
        return graph
    lower_cased = [word.lower().strip() for word in exclude_keywords]

    for line in graph:
        if len(set(line).intersection(set(lower_cased))) != 0:
            pass
        else:
            fixed_graph.append(line)
    return fixed_graph

def get_active_sheetname(xls_filepath: str) -> List[str]:
    # Converting xls file to .xlsx because openpyxl doesn't support xls
    is_temp = False
    if xls_filepath.endswith(".xls"):
        is_temp = True
        p.save_book_as(file_name=xls_filepath,
                    dest_file_name=f"{xls_filepath}x")
        xls_filepath = f"{xls_filepath}x"

    if xls_filepath.endswith(".csv"):
        is_temp = True
        p.save_book_as(file_name=xls_filepath,
                    dest_file_name=xls_filepath.replace(".csv", ".xlsx"))
        xls_filepath = xls_filepath.replace(".csv", ".xlsx")

    # Loading workbook
    workbook = openpyxl.load_workbook(xls_filepath, True)
    sheetname = workbook.active.title
    workbook.close()
    if is_temp:
        os.remove(xls_filepath)
    return sheetname

@Gooey(program_name="D2 Research Maker Toolkit",
       image_dir=resource_path("icons"),
       default_size=(1100,720),
program_description="""Simple co-occurrence analysis matrix generation tool.
Import Data from .xlsx, .xls .csv. Homogenize given data using lemmatizing.
""",
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
        optional_cols=5,
        required_cols=3,
        disable_progress_bar_animation=True
)
def main():
    parser = GooeyParser()
    parser.add_argument("filepath", metavar="Path to excel spreadsheet", type=str, widget="FileChooser", help="Choose XLS (Excel) file.",
                        gooey_options={
                            'wildcard':
                                "XLSX (Excel spreadsheet) (*.xlsx,*.xls,*.csv)|*.xlsx;*.xls;*.csv|"
                                "All files (*.*)|*.*",
                            'default_file': "Pick XLSX file",
                            'message': "Select XLSX (Excel spreadsheet) file"
                        }
                        )
    parser.add_argument("--sheet_name", metavar="Name of the sheet",help="Select the sheetname. (Leave empty to select the active spreadsheet.)")
    parser.add_argument("range", metavar="Range",type=str, help="Range of the cells that will be used in frequency analysis.\nExample: E1:E18|A6:A19, use '|' to select two ranges at once")
    parser.add_argument("--lemmatization", action='store_true', metavar="Lemmatization", widget="CheckBox", help="Groups together different inflected forms of the same word, for example:\n'tree diseases' -> 'tree disease'\n'asians' -> 'asian'", default=True)
    parser.add_argument("save_as", metavar="Save as...", help="Choose the output file name.",widget="FileSaver", default=os.path.join(get_execution_folder(),"output.xlsx"),
                        gooey_options={
                                'wildcard':
                                    "XLSX (Excel spreadsheet) (*.xlsx)|*.xlsx|"
                                    "All files (*.*)|*.*",
                                'message': "Create name for xlsx file",
                                'default_file': "output.xlsx"
                            })
    parser.add_argument("--delimeter", metavar="Delimeter for cell's data", default=";", help="Select the delimeter between keys in cell value.",)
    parser.add_argument("--exclude_keywords", type=str, metavar="Exclude specific keywords", help="If you want to remove cells that contain one of specific keywords, write them using semicolons (;) or commas (,)\nExample: Science; Climate change")
    parser.add_argument("--binary", action='store_true', metavar="Binary matrix", widget="CheckBox", help="Select if you want to make the co-occurrence matrix binary.\n(Only 0s and 1s)", default=False)

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
    Binary: {args.binary}
''')
    logger.info(f"Loading {args.filepath} file.")

    graph = load_xls_sheet_values(args.filepath, args.range)
    logger.info(f"Successfully loaded and read {args.filepath}.")

    logger.info(f"Starting to homogenize cell values.")
    graph = homogenize(graph, lemmatize_=args.lemmatization)
    logger.info("Successfully finished homogenizing cells.")

    if args.exclude_keywords:
        main_delimeter = ";"
        if len(args.exclude_keywords.split(";")) > 1:
            main_delimeter = ";"
        elif len(args.exclude_keywords.split(",")) > 1:
            main_delimeter = ","
        logger.info("Starting to exclude selected keywords.")
        graph = exclude_keywords_from_graph(graph, args.exclude_keywords.lower().split(main_delimeter))
        logger.info("Successfully excludeded selected keywords.")

    logger.info("Generating co-occurrence matrix on a homogenized cell values.")
    co_occurrence_matrix = generate_co_occurrence_matrix(graph, args.binary)
    logger.info("Successfully generated co-occurrence matrix.")

    logger.info(f"Writing to {args.save_as}")
    generate_excel(co_occurrence_matrix, args.save_as)
    logger.info("Success! The program has finished.")

if __name__ == "__main__":
    main()