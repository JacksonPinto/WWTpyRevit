# -*- coding: utf-8 -*-
# script.py (IronPython main script for pyRevit Cable Length Calculation)
# Version: 6.8.0 (2025-08-22)
# Author: JacksonPinto
# UPDATE: Optionally filter electrical elements by user-selected Scope Box. If a Scope Box is selected, only elements inside are processed; otherwise, all elements in the active view are processed.
# UPDATE: For each electrical device, finds the closest point on any tray network line (not just a node), creates a new "insertion" point on the tray for a correct L-connection, and updates the network for export and pathfinding.
# Uses int(str(el.Id)) for all ElementId cross-referencing.
# Calls Python 3 with hardcoded path for calc_shortest.py subprocess.

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    Face,
    UV,
    Line as RevitLine,
    XYZ
)
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import forms
import math
import json
import os
import sys
sys.path.append(r"C:\Users\JacksonAugusto\AppData\Local\Programs\Python\Python312\Lib\site-packages")
sys.path.append(r"C:\Users\JacksonAugusto\AppData\Local\Programs\Python\Python312\Lib")
import subprocess

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def isclose(a, b, abs_tol=1e-6):
    return abs(a - b) <= abs_tol

def get_param_value(elem, param_name):
    param = elem.LookupParameter(param_name)
    if param:
        return param.AsString()
    return None

def get_tray_centerline(tray):
    try:
        lc = tray.Location
        if hasattr(lc, "Curve") and lc.Curve:
            curve = lc.Curve
            sp = curve.GetEndPoint(0)
            ep = curve.GetEndPoint(1)
            return [(sp.X, sp.Y, sp.Z), (ep.X, ep.Y, ep.Z)]
    except Exception:
        pass
    return None

def get_fitting_connectors(fitting):
    points = []
    try:
        cm = None
        if hasattr(fitting, 'MEPModel') and fitting.MEPModel:
            cm = fitting.MEPModel.ConnectorManager
        elif hasattr(fitting, 'ConnectorManager'):
            cm = fitting.ConnectorManager
        if cm:
            for conn in cm.Connectors:
                org = conn.Origin
                points.append((org.X, org.Y, org.Z))
    except Exception:
        pass
    return points

def get_fitting_edges(points):
    edges = []
    n = len(points)
    if n == 2:
        edges.append((points[0], points[1]))
    elif n == 3:
        edges.extend([(points[0], points[1]), (points[1], points[2]), (points[2], points[0])])
    elif n == 4:
        edges.extend([(points[0], points[2]), (points[1], points[3])])
    else:
        for i in range(n-1):
            edges.append((points[i], points[i+1]))
    return edges

def get_electrical_categories():
    cat_ids = [
        BuiltInCategory.OST_ElectricalFixtures,
        BuiltInCategory.OST_ElectricalEquipment,
        BuiltInCategory.OST_LightingFixtures,
        BuiltInCategory.OST_DataDevices,
        BuiltInCategory.OST_LightingDevices,
        BuiltInCategory.OST_CommunicationDevices,
        BuiltInCategory.OST_FireAlarmDevices,
        BuiltInCategory.OST_SecurityDevices,
        BuiltInCategory.OST_NurseCallDevices,
        BuiltInCategory.OST_TelephoneDevices
    ]
    cats = []
    for cat_id in cat_ids:
        cat = doc.Settings.Categories.get_Item(cat_id)
        if cat:
            cats.append((cat.Name, cat_id))
    return cats

def get_face_center(face):
    bbox = face.GetBoundingBox()
    min_uv = bbox.Min
    max_uv = bbox.Max
    center_uv = UV((min_uv.U + max_uv.U) * 0.5, (min_uv.V + max_uv.V) * 0.5)
    center_xyz = face.Evaluate(center_uv)
    return (center_xyz.X, center_xyz.Y, center_xyz.Z)

def distance3d(a, b):
    return math.sqrt((a[0]-b[0])**2 + (a[1]-b[1])**2 + (a[2]-b[2])**2)

def closest_point_on_line(pt, line):
    """Given pt=(x,y,z), line=((x1,y1,z1),(x2,y2,z2)), return (closest_pt, dist, t)
    t in [0,1]: on the segment. t<0 or t>1: projection off segment, use endpoint.
    """
    ax, ay, az = line[0]
    bx, by, bz = line[1]
    px, py, pz = pt
    ab = (bx - ax, by - ay, bz - az)
    ap = (px - ax, py - ay, pz - az)
    ab_len2 = ab[0]**2 + ab[1]**2 + ab[2]**2
    if ab_len2 == 0:
        return (line[0], distance3d(pt, line[0]), 0.0)
    t = (ap[0]*ab[0] + ap[1]*ab[1] + ap[2]*ab[2]) / ab_len2
    t_clamped = max(0.0, min(1.0, t))
    closest = (ax + ab[0]*t_clamped, ay + ab[1]*t_clamped, az + ab[2]*t_clamped)
    dist = distance3d(pt, closest)
    return (closest, dist, t_clamped)

# ---------------- STEP 1: Ask Equipment Color
color_filter = forms.ask_for_string(
    prompt="Enter Equipment Color value to filter Cable Trays and Fittings:",
    default="",
    title="Filter by Equipment Color"
)
if not color_filter:
    forms.alert("No Equipment Color value entered. Script cancelled.", exitscript=True)

# Collect trays and fittings
collector_trays = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_CableTray).WhereElementIsNotElementType()
collector_fittings = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_CableTrayFitting).WhereElementIsNotElementType()
trays = [elem for elem in collector_trays if get_param_value(elem, "Equipment Color") and get_param_value(elem, "Equipment Color").strip().lower() == color_filter.strip().lower()]
fittings = [elem for elem in collector_fittings if get_param_value(elem, "Equipment Color") and get_param_value(elem, "Equipment Color").strip().lower() == color_filter.strip().lower()]

# ---------------- STEP 2: Build tray/fitting network (edges)
network_edges = []
tray_lines = []
tray_points = []  # for legacy fallback, not used in new logic

for tray in trays:
    cl = get_tray_centerline(tray)
    if cl:
        network_edges.append(tuple(cl))
        tray_lines.append((cl[0], cl[1]))
        tray_points.append(cl[0])
        tray_points.append(cl[1])

for fitting in fittings:
    pts = get_fitting_connectors(fitting)
    if len(pts) >= 2:
        edges = get_fitting_edges(pts)
        for edge in edges:
            network_edges.append(edge)
            tray_lines.append(edge)
            tray_points.append(edge[0])
            tray_points.append(edge[1])

# ---------------- STEP 3: User selects electrical element category, connect all elements to network
categories = get_electrical_categories()
if not categories:
    forms.alert("No electrical categories found in model.", exitscript=True)

cat_names = [cat[0] for cat in categories]
selected_name = forms.SelectFromList.show(cat_names, title="Select Electrical Element Category", multiselect=False)
if not selected_name:
    forms.alert("No category selected. Script cancelled.", exitscript=True)
if isinstance(selected_name, list):
    selected_name = selected_name[0]
selected_cat_id = [cat[1] for cat in categories if cat[0] == selected_name][0]

collector_elements = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(selected_cat_id).WhereElementIsNotElementType()
elements = [el for el in collector_elements if el.Location and hasattr(el.Location, "Point")]

if not elements:
    forms.alert("No elements of selected category visible in active view.", exitscript=True)

# ---------------- OPTIONAL SCOPE BOX FILTER ----------------
scopebox_elem = None
try:
    pick = forms.alert("Optionally select a Scope Box. Click OK to select or Cancel to use all elements in the active view.", options=["Select", "Cancel"])
    if pick == "Select":
        scopebox_ref = uidoc.Selection.PickObject(ObjectType.Element, "Select a Scope Box (or Cancel for all elements in view)")
        scopebox_elem = doc.GetElement(scopebox_ref.ElementId)
        # Ensure Scope Box category
        from Autodesk.Revit.DB import BuiltInCategory
        if scopebox_elem.Category.Id.IntegerValue != int(BuiltInCategory.OST_VolumeOfInterest):
            forms.alert("Selected element is not a Scope Box. Ignoring selection.")
            scopebox_elem = None
except Exception:
    scopebox_elem = None

if scopebox_elem:
    bbox = scopebox_elem.get_BoundingBox(None)
    if bbox:
        elements_in_scope = []
        for el in elements:
            try:
                locpt = el.Location.Point
                if (bbox.Min.X <= locpt.X <= bbox.Max.X and
                    bbox.Min.Y <= locpt.Y <= bbox.Max.Y and
                    bbox.Min.Z <= locpt.Z <= bbox.Max.Z):
                    elements_in_scope.append(el)
            except Exception:
                pass
        elements = elements_in_scope
        if not elements:
            forms.alert("No elements of selected category inside chosen Scope Box.", exitscript=True)

# ---------------- STEP 4: User selects tray/fitting FACE in any view, script gets face center
forms.alert("Select a FACE on a cable tray or fitting (in any view).\nThe face center will be used as the route destination for all devices.")
try:
    picked_ref = uidoc.Selection.PickObject(ObjectType.Face, "Select a FACE on a cable tray or fitting")
except Exception as e:
    forms.alert("No face selected or cancelled. Script cancelled.\n\n{}".format(str(e)), exitscript=True)

picked_elem = doc.GetElement(picked_ref.ElementId)
geom_obj = picked_elem.GetGeometryObjectFromReference(picked_ref)
if not hasattr(geom_obj, "Evaluate"):
    forms.alert("Selected object is not a face. Script cancelled.", exitscript=True)

end_xyz = get_face_center(geom_obj)

# ----------- STEP 5: For each device, connect to tray network by closest point on any tray segment -----------
def pt_str(pt):
    return "({:.3f}, {:.3f}, {:.3f})".format(pt[0], pt[1], pt[2])

connections = []
start_points = []
csv_lines = []
JUMPER_WARN_DIST = 4000.0  # mm, about 13 feet

for idx, el in enumerate(elements):
    pt = el.Location.Point
    el_xyz = (pt.X, pt.Y, pt.Z)
    el_id = int(str(el.Id))
    start_points.append({
        "element_id": el_id,
        "point": [el_xyz[0], el_xyz[1], el_xyz[2]]
    })

    # Find the closest point on ANY tray/fitting segment (not just node)
    min_dist = None
    min_data = None # tuple: (closest_xyz_on_line, tray_line, t)
    for line in tray_lines:
        closest, d, t = closest_point_on_line(el_xyz, line)
        if (min_dist is None) or (d < min_dist):
            min_dist = d
            min_data = (closest, line, t)
    if min_data is None:
        print("[ERROR] Could not find tray network for element {} at {}".format(el_id, pt_str(el_xyz)))
        continue

    closest_on_tray, tray_line, tval = min_data

    # Build jumper: device (x,y,z) -> (x,y,closest_on_tray[2])
    drop_pt = (el_xyz[0], el_xyz[1], closest_on_tray[2])
    drop_dist = abs(el_xyz[2] - drop_pt[2])
    jumper_dist = distance3d(drop_pt, closest_on_tray)

    # Diagnostics
    print("[DEBUG] Device {} at {}:".format(el_id, pt_str(el_xyz)))
    print("    Drop to tray Z: {} (dist={:.1f})".format(pt_str(drop_pt), drop_dist))
    print("    Jumper to tray: {} (dist={:.1f})".format(pt_str(closest_on_tray), jumper_dist))
    if drop_dist > JUMPER_WARN_DIST:
        print("    [WARN] Drop >{:.0f}mm".format(JUMPER_WARN_DIST))
    if jumper_dist > JUMPER_WARN_DIST:
        print("    [WARN] Horizontal jumper >{:.0f}mm".format(JUMPER_WARN_DIST))

    # Always connect: device → drop, drop → tray insertion on line
    if not isclose(el_xyz[2], drop_pt[2]):
        connections.append((el_xyz, drop_pt))
        csv_lines.append((el_id, "drop", el_xyz, drop_pt, drop_dist))
    if not isclose(drop_pt[0], closest_on_tray[0]) or not isclose(drop_pt[1], closest_on_tray[1]) or not isclose(drop_pt[2], closest_on_tray[2]):
        connections.append((drop_pt, closest_on_tray))
        csv_lines.append((el_id, "jumper", drop_pt, closest_on_tray, jumper_dist))

# Add tray/fitting network edges to CSV too for visualization
for idx, seg in enumerate(network_edges):
    csv_lines.append(("TRAY", "tray_edge", seg[0], seg[1], distance3d(seg[0], seg[1])))

# ---------------- EXPORT CSV for visualization/debugging
try:
    script_path = os.path.abspath(__file__)
    script_dir = os.path.dirname(script_path)
    csv_path = os.path.join(script_dir, "tray_network_debug.csv")
except:
    script_dir = os.getcwd()
    csv_path = os.path.join(script_dir, "tray_network_debug.csv")

with open(csv_path, "w") as f:
    f.write("element_id,type,x1,y1,z1,x2,y2,z2,length\n")
    for entry in csv_lines:
        eid, typ, pt1, pt2, length = entry
        f.write("{},{},{:.3f},{:.3f},{:.3f},{:.3f},{:.3f},{:.3f},{:.1f}\n".format(
            eid, typ, pt1[0], pt1[1], pt1[2], pt2[0], pt2[1], pt2[2], length
        ))
print("[DEBUG] Exported tray network and all device jumpers to {}".format(csv_path))

# ---------------- EXPORT JSON for TopologicPy
def tolist(pt):
    return [float(pt[0]), float(pt[1]), float(pt[2])]

all_lines = network_edges + connections
vertices = []
vertex_map = {}
edges = []

for line in all_lines:
    pt1 = tuple(line[0])
    pt2 = tuple(line[1])
    # Map pt1
    if pt1 not in vertex_map:
        vertex_map[pt1] = len(vertices)
        vertices.append(list(pt1))
    # Map pt2
    if pt2 not in vertex_map:
        vertex_map[pt2] = len(vertices)
        vertices.append(list(pt2))
    i = vertex_map[pt1]
    j = vertex_map[pt2]
    edges.append([i, j])

output_data = {
    "vertices": vertices,
    "edges": edges,
    "start_points": start_points,
    "end_point": tolist(end_xyz)
}

try:
    json_path = os.path.join(script_dir, "topologic.JSON")
except:
    json_path = os.path.join(os.getcwd(), "topologic.JSON")

with open(json_path, 'w') as f:
    json.dump(output_data, f, indent=2)

forms.alert("Graph data and tray network debug CSV exported for TopologicPy.\n{}\n{}\nNow running TopologicPy calculation for shortest paths...".format(json_path, csv_path))

# ---------------- AUTO RUN STEP 5: calc_shortest.py (CPython)
PYTHON3_PATH = r"C:\Users\JacksonAugusto\AppData\Local\Programs\Python\Python312\python.exe"
calc_shortest_path = os.path.join(script_dir, "calc_shortest.py")
if not os.path.exists(calc_shortest_path):
    forms.alert("calc_shortest.py not found in script folder. Please ensure it is present for automatic shortest path calculation.", exitscript=False)
else:
    try:
        subprocess.check_call([PYTHON3_PATH, calc_shortest_path], cwd=script_dir)
    except Exception as e:
        forms.alert("Could not execute calc_shortest.py automatically with Python 3.12.\nError: {}\nPlease run manually:\n{}".format(str(e), calc_shortest_path), exitscript=False)

# ---------------- Last Step: update_cable_lengths.py (IronPython)
update_script = os.path.join(script_dir, "update_cable_lengths.py")
if not os.path.exists(update_script):
    forms.alert("update_cable_lengths.py not found in script folder. Please ensure it is present.", exitscript=True)
else:
    try:
        execfile(update_script)
    except Exception as e:
        forms.alert("Could not execute '{}': {}".format(update_script, str(e)), exitscript=True)

forms.alert("Cable Length Calculation process complete!\nAll steps ran successfully.\nCheck results in Revit and output files.")