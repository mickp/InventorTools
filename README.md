# Inventor Scripts

A collection of python scripts for use with Autodesk Inventor.

These scripts use Inventor's COM interface to automate certain tasks. Use them by calling them from the command line while Inventor is running.

## drawings_to_pdf

For all open drawings, exports the drawing as a PDF and any referenced parts as STEP files.

## parts_to_step

Exports all open parts as STEP files.

## update_drawings

Modifies a named title block in all open drawings.

Iterates over all open files and modifies a named title block in drawings.

## update_iproperties

Modifies iProperties for all open files.

## update_styles

Updates the style for all open drawings.