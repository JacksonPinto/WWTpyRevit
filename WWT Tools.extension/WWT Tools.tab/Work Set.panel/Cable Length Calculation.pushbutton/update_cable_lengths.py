# -*- coding: utf-8 -*-
# update_cable_lengths.py
# Version: 6.5.1 (2025-08-22)
# Author: JacksonPinto
# IronPython/pyRevit version, updating Revit elements from Topologic results.
# Uses int(str(el.Id)) for cross-referencing element IDs.
#
# IMPROVEMENT: Also sets parameter "Equipment Color" to "Green" if 'length' exists, "Red" if not.

from Autodesk.Revit.DB import FilteredElementCollector, Transaction
from pyrevit import forms
import json
import os

doc = __revit__.ActiveUIDocument.Document

script_dir = os.path.dirname(__file__)
results_path = os.path.join(script_dir, "topologic_results.json")

with open(results_path, "r") as f:
    results = json.load(f)

# ----------------- Ask user for target element category and parameter
categories = set()
for el in FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType():
    try:
        if el.Category:
            categories.add(el.Category.Name)
    except:
        pass
categories = sorted(categories)
selected_cat = forms.SelectFromList.show(categories, title="Select Category to Update", multiselect=False)
if not selected_cat:
    forms.alert("No category selected. Script cancelled.", exitscript=True)
if isinstance(selected_cat, list):
    selected_cat = selected_cat[0]

collector = FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType()
elements = [el for el in collector if el.Category and el.Category.Name == selected_cat]

if not elements:
    forms.alert("No elements of selected category visible in active view.", exitscript=True)

# Get all parameters from first element for user to pick target
param_names = [p.Definition.Name for p in elements[0].Parameters]
target_param = forms.SelectFromList.show(param_names, title="Select Parameter to Store Cable Length", multiselect=False)
if not target_param:
    forms.alert("No parameter selected. Script cancelled.", exitscript=True)

# Ask user for the "Equipment Color" parameter (can be fixed if always same)
color_param = "Equipment Color"
if color_param not in param_names:
    # Offer to select a color parameter if not found
    color_param_opts = [n for n in param_names if "Color" in n or "colour" in n] or param_names
    color_param = forms.SelectFromList.show(color_param_opts, title="Select Color Parameter (for Green/Red)", multiselect=False)
    if not color_param:
        forms.alert("No color parameter selected. Script cancelled.", exitscript=True)

# ----------------- Update parameters
id_to_elem = {}
for el in elements:
    el_id = int(str(el.Id))  # Use int(str(el.Id)) for IronPython compatibility!
    id_to_elem[el_id] = el

updated_count = 0
color_green_count = 0
color_red_count = 0

t = Transaction(doc, "Update Cable Lengths & Equipment Color")
t.Start()
for res in results:
    el_id = res["element_id"]
    length = res.get("length")
    el = id_to_elem.get(el_id)
    if not el:
        continue

    # Set cable length if possible
    param = el.LookupParameter(target_param)
    if param and length is not None:
        try:
            param.Set(length)
            updated_count += 1
        except Exception as e:
            print("[WARN] Could not set length for element {}: {}".format(el_id, e))

    # Set Equipment Color
    color = "Green" if length is not None else "Red"
    color_p = el.LookupParameter(color_param)
    if color_p:
        try:
            color_p.Set(color)
            if color == "Green":
                color_green_count += 1
            else:
                color_red_count += 1
        except Exception as e:
            print("[WARN] Could not set color for element {}: {}".format(el_id, e))
t.Commit()

forms.alert("Cable length updated for {} elements in '{}'.\nEquipment Color set: Green={} | Red={}\nDone.".format(
    updated_count, target_param, color_green_count, color_red_count
))