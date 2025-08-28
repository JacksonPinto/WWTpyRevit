# -*- coding: utf-8 -*-
# run_script.py  (ALT experimental run-based implementation)
# Version: 0.1.0 (2025-08-27)
# Author: JacksonPinto
#
# PURPOSE:
#   Alternative test script that builds infrastructure network primarily from
#   CableTrayRun / ConduitRun (CableTrayConduitRunBase descendants) instead of
#   stitching every tray + fitting manually. Falls back to legacy element-by-element
#   extraction for:
#       - Residual trays / conduits not part of any accepted run
#       - Fittings (tray or conduit) if user chooses to include residual elements
#
# WORKFLOW (similar to main script):
#   1. User selects infrastructure categories (still offer trays/fittings/conduits/fittings).
#   2. Script gathers run objects (CableTrayRun, ConduitRun) if those categories imply tray/conduit use.
#   3. For each run:
#        * Collect its member element Ids (reflection for MemberIds / GetMemberIds).
#        * Test Equipment Color filter against member elements (ANY match = run accepted).
#        * Extract ordered path curves (reflection for GetCurves / GetPath*).
#        * Sample curves (lines pass through directly; arcs & non-lines subdivided).
#        * Add edges between successive sample points.
#   4. Optionally include residual individual infrastructure elements (centerlines + fitting connectors).
#   5. Create L-shaped jumper connections from each electrical device to nearest network segment.
#   6. Export graph as topologic.JSON (meta included) & debug CSV.
#   7. Invoke calc_shortest.py (Python 3) then update_cable_lengths.py.
#
# NOTES / LIMITATIONS:
#   - Because Revit 2025 API method names can evolve, reflection attempts multiple
#     method/property names for run path and member IDs.
#   - If *no runs* found or *no runs pass filter*, fallback network may rely entirely
#     on residual extraction (if enabled) otherwise direct (star) mode.
#   - Arcs: Sampled with chord length max ARC_SEG_LEN_FT to preserve routed length.
#   - Units: All coordinates & lengths are Revit internal feet (consistent with other scripts).
#
# COMPATIBILITY:
#   - calc_shortest.py and update_cable_lengths.py previously supplied remain compatible.
#
# ------------------------------------------------------------------------------

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    UV,
    XYZ
)
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import forms
import math, json, os, subprocess, traceback

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# ------------------------------ CONFIG ---------------------------------
ARC_SEG_LEN_FT = 2.0      # Max chord length when discretizing arcs / splines (feet)
MERGE_TOL = 1e-5          # Vertex merge tolerance (feet)
# -----------------------------------------------------------------------

# --------------------------- BASIC UTILITIES ---------------------------
def get_catid(el):
    try:
        return el.Category.Id.IntegerValue
    except:
        try:
            return int(el.Category.Id)
        except:
            return None

def get_param_value(el, pname):
    p = el.LookupParameter(pname)
    if p:
        try:
            return p.AsString()
        except:
            return None
    return None

def distance3d(a,b):
    return math.sqrt((a[0]-b[0])**2 + (a[1]-b[1])**2 + (a[2]-b[2])**2)

def isclose(a,b,t=1e-9):
    return abs(a-b)<=t

def closest_point_on_line(pt, seg):
    ax,ay,az = seg[0]; bx,by,bz=seg[1]; px,py,pz=pt
    ab=(bx-ax,by-ay,bz-az); ap=(px-ax,py-ay,pz-az)
    ab2=ab[0]**2+ab[1]**2+ab[2]**2
    if ab2==0:
        return (seg[0], distance3d(pt, seg[0]), 0.0)
    t=(ap[0]*ab[0]+ap[1]*ab[1]+ap[2]*ab[2])/ab2
    t=max(0.0,min(1.0,t))
    cx=ax+ab[0]*t; cy=ay+ab[1]*t; cz=az+ab[2]*t
    cpt=(cx,cy,cz)
    return (cpt, distance3d(pt,cpt), t)

def conduit_or_tray_centerline(el):
    try:
        loc=el.Location
        if hasattr(loc,"Curve") and loc.Curve:
            c=loc.Curve
            sp=c.GetEndPoint(0); ep=c.GetEndPoint(1)
            return [(sp.X,sp.Y,sp.Z),(ep.X,ep.Y,ep.Z)]
    except: pass
    return None

def fitting_connectors(el):
    pts=[]
    try:
        mep=getattr(el,'MEPModel',None)
        if mep:
            cm=getattr(mep,'ConnectorManager',None)
            if cm and cm.Connectors:
                for conn in cm.Connectors:
                    o=conn.Origin
                    pts.append((o.X,o.Y,o.Z))
    except: pass
    try:
        cm2=getattr(el,'ConnectorManager',None)
        if cm2 and cm2.Connectors:
            for conn in cm2.Connectors:
                o=conn.Origin
                p=(o.X,o.Y,o.Z)
                if p not in pts:
                    pts.append(p)
    except: pass
    return pts

def fitting_edges_from_connectors(points):
    edges=[]
    n=len(points)
    if n==2:
        edges.append((points[0],points[1]))
    elif n==3:
        edges += [(points[0],points[1]),(points[1],points[2]),(points[2],points[0])]
    elif n==4:
        edges += [(points[0],points[2]),(points[1],points[3])]
    else:
        for i in range(n-1):
            edges.append((points[i],points[i+1]))
    return edges

def face_center_from_reference(ref):
    el=doc.GetElement(ref.ElementId)
    geom=el.GetGeometryObjectFromReference(ref)
    if not hasattr(geom, "Evaluate"):
        return None, el
    bbox=geom.GetBoundingBox()
    cuv=UV((bbox.Min.U+bbox.Max.U)/2.0, (bbox.Min.V+bbox.Max.V)/2.0)
    xyz=geom.Evaluate(cuv)
    return (xyz.X,xyz.Y,xyz.Z), el

def get_electrical_categories():
    ids=[
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
    out=[]
    for cid in ids:
        cat=doc.Settings.Categories.get_Item(cid)
        if cat: out.append((cat.Name,cid))
    return out

# --------------------- USER INPUT ---------------------
def ask_initial():
    infra_opts=[
        ("Cable Trays", BuiltInCategory.OST_CableTray),
        ("Cable Tray Fittings", BuiltInCategory.OST_CableTrayFitting),
        ("Conduits", BuiltInCategory.OST_Conduit),
        ("Conduit Fittings", BuiltInCategory.OST_ConduitFitting),
    ]
    names=[n for n,_ in infra_opts]
    infra_sel=forms.SelectFromList.show(names, title="Select Infrastructure Categories", multiselect=True)
    if not infra_sel: forms.alert("No infrastructure categories selected.", exitscript=True)
    infra_cats=[cid for n,cid in infra_opts if n in infra_sel]

    ecs=get_electrical_categories()
    if not ecs: forms.alert("No electrical categories found.", exitscript=True)
    enames=[n for n,_ in ecs]
    elec_sel=forms.SelectFromList.show(enames, title="Select Electrical Device Categories", multiselect=True)
    if not elec_sel: forms.alert("No electrical categories selected.", exitscript=True)
    elec_ids=[cid for n,cid in ecs if n in elec_sel]

    mode=forms.SelectFromList.show(["Manual Select (Pick Devices)","All Devices in Active View"],
                                   multiselect=False,
                                   title="Electrical Element Selection Mode")
    if not mode: forms.alert("No mode selected.", exitscript=True)
    mode = "manual" if "Manual" in mode else "all"

    include_residual=forms.alert(
        "Include RESIDUAL individual elements (centerlines + fitting connectors)\nNOT part of accepted runs?",
        options=["Yes","No"])
    residual = (include_residual=="Yes")

    return infra_cats, elec_ids, elec_sel, mode, residual

infra_cats, elec_cat_ids, elec_selected_names, select_mode, include_residual = ask_initial()

color_filter=forms.ask_for_string(
    prompt="Equipment Color filter (blank=skip). Run accepted if ANY member element matches.",
    default="Blue",
    title="Color Filter"
)
if color_filter is None:
    forms.alert("Cancelled.", exitscript=True)
color_filter_norm=color_filter.strip().lower()
use_color = (color_filter_norm!="")

# -------------------- RUN COLLECTION --------------------
# We attempt to import run classes; if not present (old Revit) we gracefully fallback.
try:
    from Autodesk.Revit.DB import CableTrayRun, ConduitRun
    RUN_SUPPORT = True
except Exception:
    RUN_SUPPORT = False

run_edges=[]          # edges extracted from runs
run_curve_samples=0
run_arcs_sampled=0
accepted_run_ids=[]
skipped_run_color=0
skipped_run_empty=0
run_member_color_cache={}
run_member_counts=[]
run_errors=[]

def get_member_ids(run):
    # Try multiple options
    for attr in ("MemberIds","GetMemberIds","ElementIds","GetElementIds"):
        try:
            val=getattr(run, attr)
            if callable(val):
                res=val()
            else:
                res=val
            # Convert to list of ElementId
            try:
                return list(res)
            except:
                pass
        except:
            continue
    return []

def get_run_curves(run):
    # Attempt typical method names
    candidates=["GetCurves","GetPathCurves","GetPath","Curves","Path"]
    for name in candidates:
        try:
            attr=getattr(run,name)
            curves=attr() if callable(attr) else attr
            # Expect iterable of Curve
            return list(curves)
        except:
            continue
    return []

def sample_curve(curve):
    """Return list of XYZ triplets along curve (including endpoints). Lines -> 2 pts; arcs/splines segmented."""
    global run_arcs_sampled
    pts=[]
    try:
        # Try to detect line quickly
        if curve.GetEndPoint(0) and curve.GetEndPoint(1):
            sp=curve.GetEndPoint(0)
            ep=curve.GetEndPoint(1)
        else:
            return pts
        length=curve.Length
        # If curve is a straight line:
        curveClass=curve.GetType().Name.lower()
        if "line" in curveClass and "arc" not in curveClass:
            pts.append((sp.X,sp.Y,sp.Z))
            pts.append((ep.X,ep.Y,ep.Z))
            return pts
        # Non-line: segment based on length
        run_arcs_sampled += 1
        seg_count = max(2, int(math.ceil(length / ARC_SEG_LEN_FT)))
        for i in range(seg_count+1):
            p=curve.Evaluate((1.0/seg_count)*i, True)
            pts.append((p.X,p.Y,p.Z))
        return pts
    except:
        # fallback just endpoints
        pts.append((sp.X,sp.Y,sp.Z))
        pts.append((ep.X,ep.Y,ep.Z))
        return pts

if RUN_SUPPORT:
    # Collect runs only if Tray or Conduit categories are selected
    want_tray = any(c in infra_cats for c in [BuiltInCategory.OST_CableTray])
    want_conduit = any(c in infra_cats for c in [BuiltInCategory.OST_Conduit])
    if want_tray:
        tray_runs = list(FilteredElementCollector(doc).OfClass(CableTrayRun))
    else:
        tray_runs = []
    if want_conduit:
        conduit_runs = list(FilteredElementCollector(doc).OfClass(ConduitRun))
    else:
        conduit_runs = []
    all_runs = tray_runs + conduit_runs

    for run in all_runs:
        try:
            member_ids = get_member_ids(run)
            run_member_counts.append((int(run.Id.IntegerValue), len(member_ids)))
            members=[doc.GetElement(mid) for mid in member_ids]
            # Evaluate color
            member_colors = set()
            for m in members:
                val=get_param_value(m,"Equipment Color")
                if val: member_colors.add(val.strip().lower())
            run_member_color_cache[int(run.Id.IntegerValue)] = list(member_colors)
            if use_color:
                if not any(c == color_filter_norm for c in member_colors):
                    skipped_run_color += 1
                    continue
            # Extract curves
            curves = get_run_curves(run)
            if not curves:
                skipped_run_empty += 1
                continue
            # Build edges
            for cv in curves:
                sample_pts = sample_curve(cv)
                if len(sample_pts) < 2:
                    continue
                for i in range(len(sample_pts)-1):
                    a=sample_pts[i]; b=sample_pts[i+1]
                    if distance3d(a,b) > 1e-9:
                        run_edges.append((a,b))
                        run_curve_samples += 1
            accepted_run_ids.append(int(run.Id.IntegerValue))
        except Exception as e:
            run_errors.append("Run {} error: {}".format(run.Id, e))

# -------------------- RESIDUAL (LEGACY) EXTRACTION --------------------
residual_edges=[]
residual_fail=[]
residual_counts={"Tray":0,"Tray_edges":0,"TrayFitting":0,"TrayFitting_edges":0,
                 "Conduit":0,"Conduit_edges":0,"ConduitFitting":0,"ConduitFitting_edges":0}

if include_residual:
    # Collect explicit infra elements (Active View)
    infra_elements=[]
    for cat in infra_cats:
        col=(FilteredElementCollector(doc, doc.ActiveView.Id)
             .OfCategory(cat)
             .WhereElementIsNotElementType())
        for el in col: infra_elements.append(el)

    # Remove elements already consumed by accepted runs (if we can detect membership)
    accepted_member_ids=set()
    if RUN_SUPPORT and accepted_run_ids:
        # Flatten all member ids from accepted runs for filtering duplicates
        for run in (tray_runs+conduit_runs):
            if int(run.Id.IntegerValue) in accepted_run_ids:
                mids=get_member_ids(run)
                for mid in mids:
                    accepted_member_ids.add(int(mid.IntegerValue))
    filtered_residual=[]
    for el in infra_elements:
        if int(el.Id.IntegerValue) in accepted_member_ids:
            continue
        # color filter
        if use_color:
            val=get_param_value(el,"Equipment Color")
            if not (val and val.strip().lower()==color_filter_norm):
                continue
        filtered_residual.append(el)

    for el in filtered_residual:
        cid=get_catid(el)
        if cid==int(BuiltInCategory.OST_CableTray):
            residual_counts["Tray"]+=1
            cl=conduit_or_tray_centerline(el)
            if cl:
                residual_edges.append(tuple(cl))
                residual_counts["Tray_edges"]+=1
            else:
                residual_fail.append((int(str(el.Id)),"Tray centerline missing"))
        elif cid==int(BuiltInCategory.OST_CableTrayFitting):
            residual_counts["TrayFitting"]+=1
            pts=fitting_connectors(el)
            if len(pts)>=2:
                fes=fitting_edges_from_connectors(pts)
                residual_counts["TrayFitting_edges"]+=len(fes)
                residual_edges.extend(fes)
            else:
                residual_fail.append((int(str(el.Id)),"Tray fitting connectors<2"))
        elif cid==int(BuiltInCategory.OST_Conduit):
            residual_counts["Conduit"]+=1
            cl=conduit_or_tray_centerline(el)
            if cl:
                residual_edges.append(tuple(cl))
                residual_counts["Conduit_edges"]+=1
            else:
                residual_fail.append((int(str(el.Id)),"Conduit centerline missing"))
        elif cid==int(BuiltInCategory.OST_ConduitFitting):
            residual_counts["ConduitFitting"]+=1
            pts=fitting_connectors(el)
            if len(pts)>=2:
                fes=fitting_edges_from_connectors(pts)
                residual_counts["ConduitFitting_edges"]+=len(fes)
                residual_edges.extend(fes)
            else:
                residual_fail.append((int(str(el.Id)),"Conduit fitting connectors<2"))

# -------------------- CONSOLIDATE INFRA NETWORK --------------------
infrastructure_edges = run_edges + residual_edges
segments_for_nearest = list(infrastructure_edges)  # simple list of pair tuples

# If infrastructure empty, we will create star connections later
# -------------------- ELECTRICAL ELEMENTS --------------------
if select_mode=="manual":
    try:
        refs=uidoc.Selection.PickObjects(ObjectType.Element,
            "Pick electrical devices (categories: {})".format(", ".join(elec_selected_names)))
        picked=[doc.GetElement(r.ElementId) for r in refs]
        valid=set(elec_selected_names)
        electrical=[el for el in picked if el.Category and el.Category.Name in valid]
        if not electrical:
            forms.alert("No picked elements match selected categories.", exitscript=True)
    except Exception as e:
        forms.alert("No elements picked.\n{}".format(e), exitscript=True)
else:
    electrical=[]
    for cid in elec_cat_ids:
        electrical.extend(
            FilteredElementCollector(doc, doc.ActiveView.Id)
            .OfCategory(cid).WhereElementIsNotElementType().ToElements()
        )

electrical=[el for el in electrical if hasattr(el,"Location") and hasattr(el.Location,"Point") and el.Location.Point]
if not electrical:
    forms.alert("No electrical point-location elements found.", exitscript=True)

# Optional scope box
try:
    scope_choice=forms.alert("Scope Box filter? (Select=Pick one / Skip)", options=["Select","Skip"])
    if scope_choice=="Select":
        sb_ref=uidoc.Selection.PickObject(ObjectType.Element,"Pick Scope Box")
        sb_el=doc.GetElement(sb_ref.ElementId)
        if get_catid(sb_el)==int(BuiltInCategory.OST_VolumeOfInterest):
            bbox=sb_el.get_BoundingBox(None)
            inside=[]
            for e in electrical:
                p=e.Location.Point
                if (bbox.Min.X<=p.X<=bbox.Max.X and
                    bbox.Min.Y<=p.Y<=bbox.Max.Y and
                    bbox.Min.Z<=p.Z<=bbox.Max.Z):
                    inside.append(e)
            electrical=inside
            if not electrical:
                forms.alert("No electrical elements inside scope box.", exitscript=True)
        else:
            forms.alert("Not a scope box. Ignoring.")
except:
    pass

# -------------------- PICK END FACE --------------------
forms.alert("Select a FACE on infrastructure (tray/conduit/run member/fitting) for End Point.")
try:
    face_ref=uidoc.Selection.PickObject(ObjectType.Face,"Pick End Point Face")
except Exception as e:
    forms.alert("End Point face not selected.\n{}".format(e), exitscript=True)

end_xyz, end_elem = face_center_from_reference(face_ref)
if not end_xyz:
    forms.alert("Failed to evaluate End Point.", exitscript=True)

# Option to add end element geometry if not already present (residual mode)
if include_residual and infrastructure_edges:
    # If end element is a residual element not covered by runs and not extracted yet
    if end_elem and end_elem.Category:
        cid=get_catid(end_elem)
        need_add = False
        if cid in (int(BuiltInCategory.OST_CableTray), int(BuiltInCategory.OST_Conduit)):
            cl = conduit_or_tray_centerline(end_elem)
            if cl and (cl[0],cl[1]) not in infrastructure_edges:
                need_add=True
                infrastructure_edges.append(tuple(cl))
                segments_for_nearest.append((cl[0],cl[1]))
        elif cid in (int(BuiltInCategory.OST_CableTrayFitting), int(BuiltInCategory.OST_ConduitFitting)):
            pts=fitting_connectors(end_elem)
            if len(pts)>=2:
                edges=fitting_edges_from_connectors(pts)
                # Add only if not already present
                new_added=False
                for e in edges:
                    if e not in infrastructure_edges:
                        infrastructure_edges.append(e)
                        segments_for_nearest.append(e)
                        new_added=True
                need_add=new_added
        if need_add:
            forms.alert("End element geometry appended to infrastructure network.")

# -------------------- DEVICE CONNECTIONS (L-Jumpers) --------------------
start_points=[]
connection_edges=[]
csv_rows=[]
direct_mode = (len(segments_for_nearest)==0)

def add_csv(eid, typ, a, b):
    csv_rows.append((eid,typ,a,b,distance3d(a,b)))

for el in electrical:
    p=el.Location.Point
    xyz=(p.X,p.Y,p.Z)
    eid=int(str(el.Id))
    start_points.append({"element_id":eid,"point":[xyz[0],xyz[1],xyz[2]]})

    if not direct_mode:
        # find nearest segment
        best=None
        for seg in segments_for_nearest:
            cpt,d,t=closest_point_on_line(xyz, seg)
            if best is None or d<best[1]:
                best=(cpt,d)
        if not best:
            # fallback direct
            if xyz!=end_xyz:
                connection_edges.append((xyz,end_xyz))
                add_csv(eid,"direct_fallback",xyz,end_xyz)
            continue
        nearest=best[0]
        drop=(xyz[0],xyz[1],nearest[2])
        if not isclose(xyz[2], drop[2]):
            connection_edges.append((xyz,drop))
            add_csv(eid,"vertical_drop",xyz,drop)
        if (not isclose(drop[0],nearest[0]) or
            not isclose(drop[1],nearest[1]) or
            not isclose(drop[2],nearest[2])):
            connection_edges.append((drop,nearest))
            add_csv(eid,"horizontal_jumper",drop,nearest)
    else:
        # star
        if xyz!=end_xyz:
            connection_edges.append((xyz,end_xyz))
            add_csv(eid,"direct_no_infra",xyz,end_xyz)

# Add infrastructure edges to CSV
for e in infrastructure_edges:
    add_csv("INFRA","infra_edge", e[0], e[1])

# -------------------- EXPORT DEBUG CSV --------------------
try:
    script_path=os.path.abspath(__file__)
    script_dir=os.path.dirname(script_path)
except:
    script_dir=os.getcwd()
csv_path=os.path.join(script_dir,"run_network_debug.csv")
with open(csv_path,"w") as f:
    f.write("element_id,type,x1,y1,z1,x2,y2,z2,length\n")
    for row in csv_rows:
        i,typ,a,b,l=row
        f.write("{},{},{:.6f},{:.6f},{:.6f},{:.6f},{:.6f},{:.6f},{:.6f}\n".format(
            i,typ,a[0],a[1],a[2],b[0],b[1],b[2],l))

# -------------------- BUILD GRAPH (merge vertices) --------------------
all_edges = list(infrastructure_edges) + list(connection_edges)
vmap={}
vertices=[]
edges=[]
def add_vertex(pt):
    # Merge by tolerance
    for existing,index in zip(vertices, range(len(vertices))):
        if (abs(existing[0]-pt[0])<MERGE_TOL and
            abs(existing[1]-pt[1])<MERGE_TOL and
            abs(existing[2]-pt[2])<MERGE_TOL):
            return index
    idx=len(vertices)
    vertices.append([pt[0],pt[1],pt[2]])
    return idx

for seg in all_edges:
    i1=add_vertex(seg[0])
    i2=add_vertex(seg[1])
    if i1!=i2:
        edges.append([i1,i2])

# -------------------- META --------------------
meta={
    "version":"0.1.0-run",
    "equipment_color_filter": color_filter,
    "use_color_filter": use_color,
    "runs_supported": RUN_SUPPORT,
    "accepted_run_ids": accepted_run_ids,
    "run_edges_count": len(run_edges),
    "run_curve_samples": run_curve_samples,
    "run_arcs_sampled": run_arcs_sampled,
    "skipped_run_color": skipped_run_color,
    "skipped_run_empty": skipped_run_empty,
    "run_member_counts": run_member_counts[:30],
    "run_member_color_cache": dict(list(run_member_color_cache.items())[:30]),
    "run_errors_sample": run_errors[:10],
    "include_residual": include_residual,
    "residual_edge_count": len(residual_edges),
    "residual_counts": residual_counts,
    "residual_fail_samples": residual_fail[:10],
    "segments_built": len(infrastructure_edges),
    "direct_mode": direct_mode,
    "electrical_elements_count": len(electrical),
    "unit_notice": "All coordinates & lengths in internal feet."
}

graph_data={
    "meta": meta,
    "vertices": vertices,
    "edges": edges,
    "start_points": start_points,
    "end_point": [end_xyz[0],end_xyz[1],end_xyz[2]]
}

json_path=os.path.join(script_dir,"topologic.JSON")
with open(json_path,"w") as f:
    json.dump(graph_data,f,indent=2)

forms.alert(
    "RUN graph built.\nVertices:{} Edges:{} Devices:{}\nRunEdges:{} ResidualEdges:{} DirectMode:{}\nJSON:{}\nCSV:{}"
    .format(len(vertices), len(edges), len(start_points),
            len(run_edges), len(residual_edges), direct_mode, json_path, csv_path)
)

# -------------------- RUN calc_shortest.py --------------------
PYTHON3_PATH = r"C:\Users\JacksonAugusto\AppData\Local\Programs\Python\Python312\python.exe"
calc_script=os.path.join(script_dir,"calc_shortest.py")
if not os.path.exists(calc_script):
    forms.alert("calc_shortest.py not found (skipping computation).", exitscript=False)
else:
    try:
        subprocess.check_call([PYTHON3_PATH, calc_script], cwd=script_dir)
    except Exception as e:
        forms.alert("calc_shortest.py failed:\n{}".format(e), exitscript=False)

# -------------------- UPDATE PARAMETERS --------------------
update_script=os.path.join(script_dir,"update_cable_lengths.py")
if not os.path.exists(update_script):
    forms.alert("update_cable_lengths.py not found.", exitscript=True)
else:
    try:
        execfile(update_script)
    except Exception as e:
        forms.alert("Error executing update_cable_lengths.py:\n{}".format(e), exitscript=True)

forms.alert("Run-based Cable Length Calculation COMPLETE (run_script v0.1.0). Review results.")