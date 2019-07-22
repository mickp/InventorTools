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

"""Update the style on open drawings."""
import win32com.client as client
from win32com.client import constants
import os

# Inventor instance
inv = client.gencache.EnsureDispatch("Inventor.Application")

# Enumerate open documents.
doc = inv.ActiveDocument
path, filename = os.path.split(doc.FullFileName)

if doc.DocumentType == constants.kDrawingDocumentObject:
    doc = client.CastTo(doc, "DrawingDocument")
else:
    exit()

style = doc.StylesManager.ActiveStandardStyle
print("Active standard is %s.\n\n" % style.Name)

for style in doc.StylesManager.Styles:
    if not style.UpToDate:
        print('Style %s is out of date.' % style.Name)
        style.UpdateFromGlobal()

doc.Update()
