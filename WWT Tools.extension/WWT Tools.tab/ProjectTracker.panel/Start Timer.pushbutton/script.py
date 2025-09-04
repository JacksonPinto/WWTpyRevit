# -*- coding: utf-8 -*-
# ProjectTracker - Start Timer Button
#
# This script starts a time tracking session for the current Revit project.
# It prompts for project metadata and stores start info in pyRevit sticky storage.
from pyrevit import script, revit, forms
import time

# Get the active Revit document
doc = revit.doc
user = script.get_username()

# Prompt for project metadata (comma separated)
metadata = forms.ask_for_string(
    default="ProjectName,Discipline,Task",
    prompt="Enter project name, discipline, task (comma separated):"
)
if metadata:
    # Split and strip user input
    parts = [x.strip() for x in metadata.split(",")]
    # Ensure we always have 3 fields
    while len(parts) < 3:
        parts.append("Unknown")
    project_name, discipline, task = parts
else:
    project_name, discipline, task = "Unknown", "Unknown", "Unknown"

# Gather Revit context
context = {
    "file_name": doc.Title,
    "project_name": getattr(doc.ProjectInformation, "Name", ""),
    "view_name": getattr(doc.ActiveView, "Name", ""),
    "view_type": str(getattr(doc.ActiveView, "ViewType", "")),
}

# Store start time and metadata in pyRevit sticky
sticky = script.get_sticky("projecttracker")
sticky["start_time"] = time.time()
sticky["metadata"] = {
    "project_name": project_name,
    "discipline": discipline,
    "task": task
}
sticky["context"] = context

forms.alert(
    "Started time tracking for project '{}'.\nDon't forget to Stop when you're done!".format(project_name),
    exitscript=False
)