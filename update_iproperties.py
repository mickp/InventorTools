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

"""Update iProperties on open files.

This example corrects spacing in the author name field.
"""

import win32com.client as client
from win32com.client import constants
import os

property_sets = {
    "Inventor Summary Information": ['Title', 'Subject', 'Author', 'Revision Number'],
    "Inventor Document Summary Information": ['Category', 'Manager', 'Company'],
    "Design Tracking Properties": ['Part Number', 'Project', 'Cost Center', 'Material', 'Designer']
}

ps_by_property = dict( (val, key) for key in property_sets for val in property_sets[key])

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

    propname = "Author"
    prop = doc.PropertySets.Item(ps_by_property[propname]).Item(propname)
    value = prop.Value
    
    # TODO: put your tasks here.
    value = value.replace("MAPhillips", "M A Phillips")
    prop.Value = value
    
    doc.Update()

    print ('{:24}\t{:16}\t{}'.format(filename, prop.Value, value))