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

"""Batch process Inventor drawings to create files for the workshop.

This script will iterate over all open files in inventor. If the file
is a drawing, it will export the drawing as a PDF, and any models
referenced in the drawing as STEP files.
"""
import win32com.client as client
from win32com.client import constants
import os

# Inventor instance
inv = client.gencache.EnsureDispatch("Inventor.Application")

# Enumerate open documents.
docs = inv.Documents.VisibleDocuments

# Lists to store processed and skipped filenames.
processed = []
skipped = []

for i in range(1, docs.Count+1):
    doc = docs.Item(i)
    path, filename = os.path.split(doc.FullFileName)
    if doc.DocumentType == constants.kDrawingDocumentObject:
        # Export the drawing as a PDF.
        outname = '.'.join(filename.split('.')[0:-1] + ['pdf'])
        doc.SaveAs(os.path.join(path, outname), True)
        status = "%s\t-->  %s" % (filename, outname)
        processed.append(status)
        print(status)
    else:
        skipped.append(filename)

print("\n\n===SUMMARY===\n")
if skipped:
    print ("Skipped:")
    [print("\t", f) for f in skipped]

if processed:
    print("Processed:")
    [print("\t", f) for f in processed]
