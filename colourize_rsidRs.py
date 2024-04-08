from docx import Document
from docx.shared import RGBColor
import logging
import re
from sys import exit
import tkinter as tk
from tkinter import filedialog

# ***** Declare variables *****
color_options = {  # available colours. You can add more if you wish, but must use sequential integer for the key.
    1: ['red1', RGBColor(255, 0, 0)],
    2: ['orange1', RGBColor(255, 128, 0)],
    3: ['yellow1', RGBColor(255, 255, 0)],
    4: ['green2', RGBColor(128, 255, 0)],
    5: ['green1', RGBColor(0, 255, 0)],
    6: ['green3', RGBColor(0, 255, 128)],
    7: ['green4', RGBColor(0, 128, 255)],
    8: ['blue1', RGBColor(0, 255, 255)],
    9: ['blue2', RGBColor(0, 128, 255)],
    10: ['blue3', RGBColor(0, 0, 255)],
    11: ['purple1', RGBColor(127, 0, 255)],
    12: ['pink1', RGBColor(255, 0, 255)],
    13: ['fuchsia1', RGBColor(255, 0, 127)],
    14: ['grey1', RGBColor(128, 128, 128)],
    15: ['brown1', RGBColor(102, 0, 0)],
    16: ['brown2', RGBColor(102, 51, 0)],
    17: ['yellow2', RGBColor(102, 102, 0)],
    18: ['olive1', RGBColor(51, 102, 0)],
    19: ['green5', RGBColor(0, 102, 0)],
    20: ['green6', RGBColor(0, 102, 51)],
    21: ['green7', RGBColor(0, 102, 102)],
    22: ['blue4', RGBColor(0, 51, 102)],
    23: ['blue5', RGBColor(0, 0, 102)],
    24: ['purple2', RGBColor(51, 0, 102)],
    25: ['purple3', RGBColor(102, 0, 102)],
    26: ['pink2', RGBColor(102, 0, 51)],
    27: ['red2', RGBColor(255, 102, 102)],
    28: ['orange2', RGBColor(255, 178, 102)],
    29: ['yellow3', RGBColor(255, 255, 102)],
    30: ['olive2', RGBColor(178, 255, 102)],
    31: ['green8', RGBColor(102, 255, 102)],
    32: ['green9', RGBColor(102, 255, 178)],
    33: ['blue6', RGBColor(102, 255, 255)],
    34: ['blue7', RGBColor(102, 178, 255)],
    35: ['blue8', RGBColor(102, 102, 255)],
    36: ['purple4', RGBColor(178, 102, 255)],
    37: ['fuchsia2', RGBColor(255, 102, 255)],
    38: ['pink3', RGBColor(255, 102, 178)],
    39: ['grey2', RGBColor(192, 192, 192)]
}


def setup_logger(outputlog):
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


def apply_colour(rsids, excludedrsids):
    # This associates colours to RSID values
    global color_options
    colour_count = len(color_options.keys())  # determine how many colours are available
    rsid_count = len(rsids.keys())

    if rsid_count > colour_count:
        print(f'There are {rsid_count} rsidR values, but only {colour_count} colours available.')
        print('Excess count of RSIDs will not be colourized.')

        logger.warning(f'There are {rsid_count} rsidR values, but only {colour_count} colours available.')
        logger.warning('Excess count of RSIDs will not be colourized.')

    r_count = 1

    for rsid in rsids.keys():
        if r_count <= colour_count:
            rsids[rsid] = color_options[r_count][1]  # assign a colour
            logger.info(f'rsid {rsid} assigned colour {color_options[r_count][1]} ({color_options[r_count][0]})')
        r_count += 1

    return rsids, excludedrsids


if __name__ == "__main__":
    input_document = document_to_process()  # select docx file to process

    file_path = input_document[0:input_document.rindex("/") + 1]  # extract the path from the selected file.

    output_document = output_file(file_path)  # prompt for an output file (open Explorer at path of input file)
    logger = setup_logger(output_document)

    rsidFile, rsidValues, excludedValues, invalidValues = list_of_rsidr(file_path)  # select rsidr text file.

    rsidValues, excludedValues = apply_colour(rsidValues, excludedValues)  # apply colour scheme to rsidR values

    color_code_by_rsid(input_document, output_document, rsidValues)  # colourize the document, saving to the output file

    logger.info(f'rsidR file: {rsidFile}')
    logger.info(f'rsid values: {rsidValues}')
    logger.warning(f'rsid values excluded: {excludedValues}')
    logger.warning(f'Invalid values in rsid text file: {invalidValues}')
    logger.info(f'Input file: {input_document}')
    logger.info(f'Output file: {output_document}')
