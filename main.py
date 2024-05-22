# Copyright (C) 2024 Beksultan Artykbaev - All Rights Reserved

import os
import logging
from uuid import uuid4
from collections import Counter

from gooey import GooeyParser, Gooey

from core import generate_co_occurrence_matrix, exclude_keywords_from_graph, lemmatize, filter_by_frequency, homogenize
from path_utils import resource_path, get_execution_folder
from spreadsheet import generate_excel, load_xls_sheet_values, read_savedrecs
from download_lemmatizers import models

__version__ = "0.1.1"
APP_NAME = "Artyk - Research Analyser"

# Setting up logger
logging.basicConfig(level=logging.INFO, format='%(asctime)s - [%(filename)s:%(lineno)d]: %(message)s')
logger = logging.getLogger(__name__)

@Gooey(program_name=APP_NAME,
       image_dir=resource_path("icons"),
       default_size=(1100,790),
program_description="""Simple co-occurrence analysis matrix generation tool.
Import Data from .xlsx, .xls .csv, tabdelimited file (.txt), Homogenize given data using lemmatizing.
""",
       menu=[{
        "name": "Help",
        "items": [{
                "type": "AboutDialog",
                "menuTitle": "About",
                "name": APP_NAME,
                "description": "Research toolkit for universities all around the world.",
                "version": __version__,
                "copyright": "Copyright (C) 2024 Beksultan Artykbaev - All Rights Reserved",
                "website": "mailto:artykbaev2003@gmail.com",
            }]
        },
        ],
        optional_cols=3,
        required_cols=3,
        disable_progress_bar_animation=True,
        tabbed_groups=True,
        advanced=True,
)
def main() -> None:
    parser = GooeyParser()
    subs = parser.add_subparsers(help='commands', dest='command')

    co_occurrence_parser = subs.add_parser('co-occurrence-analysis', help='''Simple co-occurrence analysis matrix generation tool.
Import Data from .xlsx, .xls .csv. Homogenize given data using lemmatizing.''')
    co_occurrence_parser.add_argument("filepaths", metavar="Path(es) to excel spreadsheet(s).", nargs='+', type=str, widget="MultiFileChooser",
                        help="Choose path(es) to spreadsheet file.\n(.xlsx, .xls, .csv) or to tabdelimited file (savedrecs.txt)",
                        gooey_options={
                            'wildcard':
                                "XLSX (Excel spreadsheet) (*.xlsx,*.xls,*.csv,*.txt)|*.xlsx;*.xls;*.csv;*.txt|"
                                "All files (*.*)|*.*",
                            'default_file': "Pick XLSX or tab delimited file",
                            'message': "Select XLSX (Excel spreadsheet) file"
                        }
                        )
    co_occurrence_parser.add_argument("--sheet_name", metavar="Name of the sheet",help="Select the sheetname. (Leave empty to select the active spreadsheet.)")
    co_occurrence_parser.add_argument("range", metavar="Range",type=str, help="Range of the cells that will be used in frequency analysis.\nExample: E1:E18|A6:A19, use '|' to select two ranges at once")
    co_occurrence_parser.add_argument("--lemmatize", action='store_true', metavar="Lemmatization", widget="CheckBox", help="Groups together different inflected forms of the same word, for example:\n'tree diseases' -> 'tree disease'\n'asians' -> 'asian'")
    co_occurrence_parser.add_argument("--lemmatization_language", metavar="Lemmatization language", help="Choose the language of your document.\n(Lemmatization for this language will be applied).",widget="Dropdown", choices=[model[0].upper()+model[1::] for model in models], default="English")
    co_occurrence_parser.add_argument("save_as", metavar="Save as...", help="Choose the output file name.",widget="FileSaver",
                        default=os.path.join(get_execution_folder(),"output.xlsx"),
                        gooey_options={
                                'wildcard':
                                    "XLSX (Excel spreadsheet) (*.xlsx)|*.xlsx|"
                                    "All files (*.*)|*.*",
                                'message': "Create a name for the xlsx file",
                            })
    co_occurrence_parser.add_argument("--delimeter", metavar="Delimeter for cell's data", default=";", help="Select the delimeter between keys in cell value.\nFor your original document.",)
    co_occurrence_parser.add_argument("--exclude_keywords", type=str, metavar="Exclude specific keywords", help="If you want to remove cells that contain one of specific keywords, write them using semicolons (;) or commas (,)\nExample: Science; Climate change")
    co_occurrence_parser.add_argument("--binary", action='store_true', metavar="Binary matrix", widget="CheckBox", help="Select if you want to make the co-occurrence matrix binary.\n(Only 0s and 1s)", default=False)
    co_occurrence_parser.add_argument("--homogenize", action='store_true', metavar="Convert to lower case", widget="CheckBox", help="Select if you want to convert data in cells to lower cased version.", default=False)
    co_occurrence_parser.add_argument("--filter", metavar="Filterings", help="Reduce number of keywords to the given value, uses keyword frequency to filter.\nSignificantly speeds up calculating process.\nSet to 0 to disable.", widget="IntegerField", required=False, default=0)
    co_occurrence_parser.add_argument("--frequency", action='store_true', metavar="Frequency Analysis", widget="CheckBox", help="Select if you want to add sheet with frequency analysis.", default=True)
    args = parser.parse_args()

    if args.command == "co-occurrence-analysis":
        logger.info("Starting algorithm.")
        logger.info(f'''The settings are:
    ----------------------------------------------------
                {APP_NAME} {__version__}    
    ----------------------------------------------------
        Excel spreadsheet path: {args.filepaths}
        Sheet name: {"active" if args.sheet_name == None else args.sheet_name}
        Range: {args.range}
        Lemmatization: {args.lemmatize}
        Lemmatization language: {args.lemmatization_language}
        Save As: {args.save_as}
        Delimeter: {repr(args.delimeter)}
        Keywords to exclude: {args.exclude_keywords}
        Binary: {args.binary}
        Convert to lower case: {args.homogenize}
        Filter (Leave only): {args.filter if args.filter != 0 else "All keywords"}
        Frequency analysis: {args.frequency}
    ----------------------------------------------------''')
        logger.info(f"Loading {args.filepaths} file.")

        graph = list()
        for filepath in args.filepaths:
            if filepath.endswith(".txt"):
                temp_name = filepath.split(".")[0] + str(uuid4()) + ".xlsx"
                generate_excel(read_savedrecs(filepath), temp_name)
                for element in load_xls_sheet_values(temp_name, args.range, args.sheet_name, args.delimeter):
                    graph.append(element)
                os.remove(temp_name)
            else:
                for element in load_xls_sheet_values(filepath, args.range, args.sheet_name, args.delimeter):
                    graph.append(element)
        logger.info(f"Successfully loaded and read {args.filepaths}.")

        # Counting frequency
        most_common = None
        if args.frequency:
            logger.info("Calculating frequency.")
            all_keywords = list()
            for line in graph:
                for keyword in line:
                    all_keywords.append(keyword)
            
            # Sorting by quantity
            c = Counter(all_keywords)
            most_common = c.most_common()
            logger.info("Finished calculating frequency.")

        if int(args.filter) != 0:
            logger.info(f"Starting to filter down to {args.filter} keywords.")
            graph = filter_by_frequency(graph, int(args.filter))
            logger.info("Finished filtering.")

        if args.exclude_keywords:
            logger.info("Starting to exclude selected keywords.")
            main_delimeter = ";"
            if len(args.exclude_keywords.split(";")) > 1:
                main_delimeter = ";"
            elif len(args.exclude_keywords.split(",")) > 1:
                main_delimeter = ","
            graph = exclude_keywords_from_graph(graph, args.exclude_keywords.lower().split(main_delimeter))
            logger.info("Successfully excludeded selected keywords.")

        if args.homogenize:
            logger.info(f"Starting to homogenizing (converting to lower case) cell values.")
            graph = homogenize(graph)
            logger.info("Successfully finished homogenizing cells.")

        if args.lemmatize:
            logger.info(f"Starting to lemmatize cell values.")
            graph = lemmatize(graph, language=args.lemmatization_language)
            logger.info("Successfully finished lemmatizing cells.")

        logger.info("Generating co-occurrence matrix.")
        co_occurrence_matrix = generate_co_occurrence_matrix(graph, args.binary)
        logger.info("Successfully generated co-occurrence matrix.")

        logger.info(f"Writing to {args.save_as}")

        generate_excel(co_occurrence_matrix, output_filename=args.save_as, frequency_analysis=most_common)
        logger.info("Success! The program has finished.")


if __name__ == "__main__":
    main()