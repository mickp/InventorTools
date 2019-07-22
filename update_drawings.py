#!/usr/bin/python
# -*- coding: utf-8
#
# Copyright 2017 Mick Phillips (mick.phillips@gmail.com)
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.

"""Upate the title block on open inventor drawings.

Includes examples to correct author name in a title block
named "MAP". """
import win32com.client as client
from win32com.client import constants
from itertools import chain
import os

# Mapping and inverse mapping of property sets to properties.
property_sets = {
    "Inventor Summary Information": ['Title', 'Subject', 'Author', 'Revision Number'],
    "Inventor Document Summary Information": ['Category', 'Manager', 'Company'],
    "Design Tracking Properties": ['Part Number', 'Project', 'Cost Center', 'Material', 'Designer']
}
ps_by_property = dict( (val, key) for key in property_sets for val in property_sets[key])

# TODO: set your title block name here
# Title block to set.
titleblock_name = 'MAP'

# Inventor instance
inv = client.gencache.EnsureDispatch("Inventor.Application")

# Enumerate open documents.
docs = inv.Documents.VisibleDocuments

# Lists to store processed and skipped filenames.
processed = []
skipped = []
styles_status = []

for i in range(1, docs.Count+1):
    doc = docs.Item(i)
    path, filename = os.path.split(doc.FullFileName)

    if doc.DocumentType == constants.kDrawingDocumentObject:
        doc = client.CastTo(doc, "DrawingDocument")
    else:
        continue

    # TODO: put your tasks here.    
    # Update author property/
    propname = "Author"
    prop = doc.PropertySets.Item(ps_by_property[propname]).Item(propname)
    value = prop.Value
    value = value.replace("MAPhillips", "M A Phillips")
    value = value.replace("map", "M A Phillips")
    prop.Value = value

    
    # Set titleblock to that specified above and purge others.
    titleblock = None
    try:
        titleblock = doc.TitleBlockDefinitions.Item(titleblock_name)
    except:
        print('Title block "%s" not found in %s.' % (titleblock_name, filename) )
    
    if titleblock:
        for sheet in doc.Sheets:
            try:
                sheet.AddTitleBlock(titleblock)
            except:
                pass

        for sheet in doc.SheetFormats:
            if sheet.ReferencedTitleBlockDefinition.Name != titleblock_name:
                sheet.Delete()

        for block in doc.TitleBlockDefinitions:
            if block.Name != titleblock_name:
                block.Delete()

    # Check if the style is up to date.
    style = doc.StylesManager.ActiveStandardStyle
    if not style.UpToDate:
        styles_status.append((filename, style.Name))

    doc.Update()
    # doc.Save()
