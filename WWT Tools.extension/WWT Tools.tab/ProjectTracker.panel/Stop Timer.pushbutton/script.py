# -*- coding: utf-8 -*-
# ProjectTracker - Stop Timer Button
#
# This script stops the current time tracking session and logs the elapsed time to CSV.
from pyrevit import script, forms
import time
import os
import csv

# Retrieve sticky storage
sticky = script.get_sticky("projecttracker")

# Check if a timer was started
if "start_time" not in sticky:
    forms.alert("No tracking session in progress!", exitscript=True)

start_time = sticky["start_time"]
metadata = sticky["metadata"]
context = sticky["context"]

elapsed = int(time.time() - start_time)
user = script.get_username()

# Prepare log data
log_data = {
    "datetime": time.strftime("%Y-%m-%d %H:%M:%S"),
    "username": user,
    "project_name": metadata["project_name"],
    "discipline": metadata["discipline"],
    "task": metadata["task"],
    "file_name": context["file_name"],
    "view_name": context["view_name"],
    "view_type": context["view_type"],
    "log_type": "Automatic",
    "elapsed_seconds": elapsed
}

# Set log file path (you may want to customize this location)
data_dir = os.path.join(os.path.dirname(__file__), "..", "..", "..", "data")
if not os.path.exists(data_dir):
    os.makedirs(data_dir)
csv_path = os.path.abspath(os.path.join(data_dir, "project_log.csv"))

# Write to CSV
file_exists = os.path.isfile(csv_path)
with open(csv_path, "a", newline="") as f:
    writer = csv.DictWriter(f, fieldnames=[
        "datetime", "username", "project_name", "discipline", "task",
        "file_name", "view_name", "view_type", "log_type", "elapsed_seconds"
    ])
    if not file_exists:
        writer.writeheader()
    writer.writerow(log_data)

# Clean up sticky storage
del sticky["start_time"]
del sticky["metadata"]
del sticky["context"]

forms.alert("Tracking stopped.\nLogged {} seconds for project '{}'.".format(elapsed, metadata["project_name"]))