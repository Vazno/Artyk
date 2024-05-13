# Copyright (C) 2024 Beksultan Artykbaev - All Rights Reserved

import os
import logging

from gooey import GooeyParser, Gooey

from core import generate_co_occurrence_matrix, exclude_keywords_from_graph, homogenize, filter_by_frequency
from path_utils import resource_path, get_execution_folder
from spreadsheet import get_active_sheetname, generate_excel, load_xls_sheet_values
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
Import Data from .xlsx, .xls .csv. Homogenize given data using lemmatizing.
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
        tabbed_groups=True
)
def main() -> None:
    parser = GooeyParser()
    parser.add_argument("filepaths", metavar="Path(es) to excel spreadsheet(s)", nargs='+', type=str, widget="MultiFileChooser",
                        help="Choose path(es) to spreadsheet file.\n(.xlsx, .xls, .csv)",
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
    parser.add_argument("--lemmatization_language", metavar="Lemmatization language", help="Choose the language of your document.\n(Lemmatization for this language will be applied).",widget="Dropdown", choices=[model[0].upper()+model[1::] for model in models], default="English")
    parser.add_argument("save_as", metavar="Save as...", help="Choose the output file name.",widget="FileSaver",
                        default=os.path.join(get_execution_folder(),"output.xlsx"),
                        gooey_options={
                                'wildcard':
                                    "XLSX (Excel spreadsheet) (*.xlsx)|*.xlsx|"
                                    "All files (*.*)|*.*",
                                'message': "Create a name for the xlsx file",
                            })
    parser.add_argument("--delimeter", metavar="Delimeter for cell's data", default=";", help="Select the delimeter between keys in cell value.\nFor your original document.",)
    parser.add_argument("--exclude_keywords", type=str, metavar="Exclude specific keywords", help="If you want to remove cells that contain one of specific keywords, write them using semicolons (;) or commas (,)\nExample: Science; Climate change")
    parser.add_argument("--binary", action='store_true', metavar="Binary matrix", widget="CheckBox", help="Select if you want to make the co-occurrence matrix binary.\n(Only 0s and 1s)", default=False)
    parser.add_argument("--filter", metavar="Filterings", help="Reduce number of keywords to the given value, uses keyword frequency to filter.\nSignificantly speeds up calculating process.\nSet to 0 to disable.", widget="IntegerField", required=False, default=0)

    args = parser.parse_args()
    logger.info("Starting algorithm.")
    logger.info(f'''The settings are:
----------------------------------------------------
            {APP_NAME} {__version__}    
----------------------------------------------------
    Excel spreadsheet path: {args.filepaths}
    Sheet name: {"active" if args.sheet_name == None else args.sheet_name}
    Range: {args.range}
    Lemmatization: {args.lemmatization}
    Lemmatization language: {args.lemmatization_language}
    Save As: {args.save_as}
    Delimeter: {repr(args.delimeter)}
    Keywords to exclude: {args.exclude_keywords}
    Binary: {args.binary}
    Filter (Leave only): {args.filter if args.filter != 0 else "All keywords"}
----------------------------------------------------''')
    logger.info(f"Loading {args.filepaths} file.")

    graph = list()
    for filepath in args.filepaths:
        # Loading values from each spreadsheet file
        for element in load_xls_sheet_values(filepath, args.range, args.sheet_name, args.delimeter):
            graph.append(element)

    logger.info(f"Successfully loaded and read {args.filepaths}.")

    if int(args.filter) != 0:
        logger.info(f"Starting to filter down to {args.filter} keywords.")
        graph = filter_by_frequency(graph, int(args.filter))
        logger.info("Finished filtering.")

    logger.info(f"Starting to homogenize cell values.")
    graph = homogenize(graph, lemmatize_=args.lemmatization, language=args.lemmatization_language)
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