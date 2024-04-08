from docx import Document
from docx.shared import RGBColor
import json
import logging
import re
from sys import exit
import tkinter as tk
from tkinter import filedialog


# ***** Declare variables *****
def setup_logger(outputlog):
    global logger
    # Create a logger
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.DEBUG)

    # Create a file handler for the logger
    handler = logging.FileHandler(outputlog + '.log')
    handler.setLevel(logging.DEBUG)

    # Create a formatter and set it for the handler
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)

    # Add the handler to the logger
    logger.addHandler(handler)

    return logger


def color_code_by_rsid(docx_path, output_path, color_scheme):
    global logger
    # Load the Word document
    doc = Document(docx_path)

    # Iterate through paragraphs and runs, applying color based on RSID
    for paragraph in doc.paragraphs:
        rsidRDefault = paragraph._element.get(
            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidRDefault')
        for run in paragraph.runs:
            rsid = run.element.get(
                '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidR')  # Get RSID from run element
            if rsid is None:
                rsid = rsidRDefault  # will need to get the default RSID for that paragraph.
            if rsid in color_scheme and color_scheme[rsid] != "":  # the RSID exists, and it's colour is not blank.
                run.font.color.rgb = color_scheme[rsid]
                logger.info(f'rsid {rsid} colourised using colour {color_scheme[rsid]}')
            else:
                if rsid in color_scheme and color_scheme[rsid] == "":
                    logger.warning(f'rsid: {rsid} has no assigned colour.')
                elif rsid not in color_scheme:
                    logger.warning(f'rsid: {rsid} not found in in colour scheme.')

    # Save the modified document
    doc.save(output_path)


def document_to_process():
    # Prompts user to select the input document.

    root = tk.Tk()
    root.withdraw()  # Hide the main window

    msword_file = filedialog.askopenfilename(title="Select DOCx file to process", initialdir=".",
                                             filetypes=[("DOCx File", "*.docx")])  # ask for file to process
    if not msword_file:  # if no docx file selected
        print(f'No DOCx file selected.')
        wait = input(f'Press ENTER to exit application...')
        exit()

    return msword_file


def list_of_rsidr(fpath):
    # prompts for the text file with rsidR values.
    global logger

    rsidr_values = {}  # empty dictionary to add rsidR values in the text file as keys.
    excluded_rsidr_values = []  # empty list to contain rsidR values commented out.
    invalid_entries = []  # empty list to contain invalid entries in rsidR file.

    root = tk.Tk()
    root.withdraw()  # Hide the main window

    rsidr_file = filedialog.askopenfilename(title="Select text file containing list of rsidR values (one per line)",
                                            initialdir=fpath)  # ask for file to process

    if not rsidr_file:  # if no docx file selected
        print(f'No file selected.')
        wait = input(f'Press ENTER to exit application...')
        exit()

    logger.info(f'Processing rsid values in {rsidr_file}.')
    with open(rsidr_file, 'r') as rfile:
        for rsid in rfile:
            if re.match(r'^[0-9a-fA-F]{8}$', rsid.strip()):  # 8 hex characters - rsid pattern
                rsidr_values[rsid.strip()] = ""  # add it to the dictionary
                logger.info(f'rsid {rsid.strip()} validated.')
            elif re.match(r'^#.{0,5}[0-9a-fA-F]{8}$', rsid.strip()):  # Valid rsid commented out
                excluded_rsidr_values.append(rsid.strip())  # add it to list of excluded rsid values
                logger.info(f'rsid {rsid.strip()} commented out in text file. Ignored.')
            elif rsid.strip() == "":  # empty line
                pass  # do nothing with it
            else:  # something else
                invalid_entries.append(rsid.strip())
                logger.warning(f'Invalid value: {rsid.strip()}')

    return rsidr_file, rsidr_values, excluded_rsidr_values, invalid_entries


def output_file(fpath):
    # Prompts user to select the output document.
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    msword_output_file = filedialog.asksaveasfilename(title="Output file (new file)", initialdir=fpath,
                                                      filetypes=[("DOCx File", "*.docx")])  # ask for file to process

    if not msword_output_file:  # if no docx file selected
        print(f'No output file provided.')
        wait = input(f'Press ENTER to exit application...')
        exit()

    if msword_output_file[-5:] != ".docx":  # checks if the file name given has the extension .docx
        msword_output_file = msword_output_file + ".docx"  # if no, add it.

    return msword_output_file


def colour_file(cpath):
    # Prompts user to select the colour file in JSON format.
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    cfile = filedialog.askopenfilename(title="Select colour file", initialdir=cpath,
                                               filetypes=[("json File", "*.json")])  # ask for file to process
    if not cfile:  # if no docx file selected
        print(f'No colour file selected.')
        wait = input(f'Press ENTER to exit application...')
        exit()

    with open('colours.json') as colours:
        c_options = json.load(colours)

    return cfile, c_options


def apply_colour(rsids, excludedrsids):
    # This associates colours to RSID values
    global colour_options, logger
    colour_count = len(colour_options.keys())  # determine how many colours are available
    rsid_count = len(rsids.keys())

    if rsid_count > colour_count:
        print(f'There are {rsid_count} rsidR values, but only {colour_count} colours available.')
        print('Excess count of RSIDs will not be colourized.')

        logger.warning(f'There are {rsid_count} rsidR values, but only {colour_count} colours available.')
        logger.warning('Excess count of RSIDs will not be colourized.')

    r_count = 1

    for rsid in rsids.keys():
        if r_count <= colour_count:
            k = str(r_count)  # convert to string, as keys are strings in JSON file.
            rsids[rsid] = RGBColor(colour_options[k][1][0], colour_options[k][1][1],
                                   colour_options[k][1][2])  # assign a colour
            logger.info(f'rsid {rsid} assigned colour {colour_options[k][1]} ({colour_options[k][0]})')
        r_count += 1

    return rsids, excludedrsids


if __name__ == "__main__":
    input_document = document_to_process()  # select docx file to process

    file_path = input_document[0:input_document.rindex("/") + 1]  # extract the path from the selected file.

    output_document = output_file(file_path)  # prompt for an output file (open Explorer at path of input file)

    logger = setup_logger(output_document)
    logger.info(f'Input file: {input_document}')

    rsidFile, rsidValues, excludedValues, invalidValues = list_of_rsidr(file_path)  # select rsidr text file.

    colourFile, colour_options = colour_file(file_path)

    rsidValues, excludedValues = apply_colour(rsidValues, excludedValues)  # apply colour scheme to rsidR values

    color_code_by_rsid(input_document, output_document, rsidValues)  # colourize the document, saving to the output file

    logger.info(f'rsid values: {rsidValues}')
    logger.warning(f'rsid values excluded: {excludedValues}')
    logger.warning(f'Invalid values in rsid text file: {invalidValues}')
    logger.info(f'Colour file: {colourFile}.')
    logger.info(f'Output file: {output_document}')
