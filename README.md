# MS Word - Colourize text based on rsidR values

**Version 1.4**

## VALIDATE YOUR FINDINGS

This application assists with identifying visible text associated with Microsoft Word `w:rsidR` revision-save identifiers. Always validate findings you intend to rely on by examining the underlying OOXML, especially `word/document.xml`. RSID analysis is corroborative and should not be treated as conclusive authorship attribution.

## Overview

Open a `.docx` file and the application identifies active `rsidR` values, lists them in the left pane, and shows their occurrence counts in descending order. Select the colour swatch or rsidR label to assign a colour. The document preview updates immediately.

The application uses an **effective rsidR** rule for visible text: a run's own `w:rsidR` takes precedence; if the run does not carry one, the containing paragraph's `w:rsidR` is used. Version 1.4 applies this same rule to both the preview/report and saved colourized document.

## rsidR pane, scrolling, and filtering

The rsidR pane includes a persistent vertical scrollbar. Use the scrollbar or mouse wheel to move through all active values. The filter box can also locate a specific rsidR value; filtering does not remove or change colour assignments.

## Saving a colourized version

Use **Save** or **Save As...** to create a colourized copy. Version 1.4 corrects an issue in v1.3 where text inheriting its rsidR from the containing paragraph could appear coloured in the preview/report but fail to receive that colour in the saved DOCX.

A corresponding colourization log is created. It records source/output hashes, original metadata, assigned rsidR values, and the associated hex/RGB colours.

## rsidR Report

Use **Generate Report** to create a Word report mapping each active rsidR value to page number, paragraph number, and extracted text segments. Page-number resolution uses Microsoft Word automation through `pywin32`; when Word automation is unavailable, page numbers may not be populated.

The recommended workflow is:

1. Open the original document.
2. Generate and review the RSID report.
3. Filter for rsidR values of interest.
4. Assign colours.
5. Save a colourized copy.
6. Validate important findings against `word/document.xml`.

## Logging

Opening a document creates an examination log. Saving a colourized copy creates a colourization log. Retain these logs with the analytical work product where reproducibility is required.

## Dependencies

- Python 3.11 or later
- `python-docx>=1.1.0`
- `pywin32>=306` on Windows for Microsoft Word page-number automation
- Tk/Tkinter for the graphical interface

Install Python dependencies with:

```text
pip install -r requirements.txt
```

On Debian/Ubuntu, Tkinter may need to be installed separately:

```text
sudo apt-get install python3-tk
```

## Version 1.4 changes

- Fixed saved colourization for runs that inherit `w:rsidR` from their containing paragraph.
- Made save-time rsidR resolution consistent with the preview and RSID report.
- Corrected the rsidR pane layout so the vertical scrollbar remains visible beside the expanding canvas.
- Documentation updated to describe effective rsidR handling and scrolling.

## Executable version

If a packaged `.exe` is provided, it can be used instead of running the Python script directly.
