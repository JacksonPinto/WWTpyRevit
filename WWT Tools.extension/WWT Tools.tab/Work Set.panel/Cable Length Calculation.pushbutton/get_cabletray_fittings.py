# Dynamo Python: Get all connector points from Cable Tray Fittings
# Works in CPython3 and IronPython2 engines.
#
# Inputs:
#   IN[0] : A Cable Tray Fitting element, an ElementId/int, or a list (nested ok) of them.
#   IN[1] : as_dynamo_points (bool, optional, default True). If True -> Dynamo Points, else -> XYZ tuples (feet).
#   IN[2] : filter_to_cabletray_domain (bool, optional, default True). Keep only cable tray/conduit connectors.
#
# Output:
#   OUT   : List of connector points per input element (structure matches input).

import clr

# RevitServices
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
doc = DocumentManager.Instance.CurrentDBDocument

# Revit API
clr.AddReference('RevitAPI')
from Autodesk.Revit import DB

# RevitNodes conversion (for .ToPoint())
clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.GeometryConversion)

# ProtoGeometry (fallback)
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import Point as DSPoint

# --------------------------
# Unwrap helpers (DB.Element)
# --------------------------
def _unwrap(elem):
    if elem is None:
        return None
    # Accept ElementId or plain int
    if isinstance(elem, DB.ElementId):
        return doc.GetElement(elem)
    if isinstance(elem, int):
        try:
            return doc.GetElement(DB.ElementId(int(elem)))
        except Exception:
            return None
    # Dynamo wrapper element may expose .InternalElement
    try:
        ie = getattr(elem, "InternalElement", None)
        if ie:
            return ie
    except Exception:
        pass
    # Dynamo global UnwrapElement (provided by engine)
    try:
        return UnwrapElement(elem)
    except Exception:
        pass
    # If it's already a DB.Element return as is; else give up
    return elem if isinstance(elem, DB.Element) else None

def _as_ds_point(xyz):
    try:
        return xyz.ToPoint()
    except Exception:
        return DSPoint.ByCoordinates(float(xyz.X), float(xyz.Y), float(xyz.Z))

def _xyz_tuple(xyz):
    return (float(xyz.X), float(xyz.Y), float(xyz.Z))

def _is_ct_fitting(db_elem):
    try:
        if db_elem is None or db_elem.Category is None:
            return False
        return int(db_elem.Category.Id.IntegerValue) == int(DB.BuiltInCategory.OST_CableTrayFitting)
    except Exception:
        return False

def _is_ct_domain(conn):
    # Keep CableTray/Conduit domain
    try:
        dom = conn.Domain
        # Exact enum compare where available
        try:
            if dom == getattr(DB.Domain, "DomainCableTrayConduit"):
                return True
        except Exception:
            pass
        s = str(dom).lower()
        return ("cabletray" in s) or ("conduit" in s)
    except Exception:
        return True  # if unknown, don't discard

def _collect_connectors(db_elem, filter_ct_domain=True):
    conns = []

    if db_elem is None:
        return conns

    # FamilyInstance MEPModel -> ConnectorManager
    try:
        mep = getattr(db_elem, "MEPModel", None)
        if mep:
            cm = getattr(mep, "ConnectorManager", None)
            if cm and cm.Connectors:
                for c in cm.Connectors:
                    conns.append(c)
    except Exception:
        pass

    # Some types expose ConnectorManager directly
    try:
        cm2 = getattr(db_elem, "ConnectorManager", None)
        if cm2 and cm2.Connectors:
            for c in cm2.Connectors:
                conns.append(c)
    except Exception:
        pass

    # Deduplicate
    uniq, seen = [], set()
    for c in conns:
        try:
            key = (getattr(c, "Owner", None).Id.IntegerValue, getattr(c, "Id", None))
        except Exception:
            key = repr(c)
        if key not in seen:
            seen.add(key)
            uniq.append(c)

    if filter_ct_domain:
        uniq = [c for c in uniq if _is_ct_domain(c)]

    return uniq

def _get_points_for_elem(elem, as_dynamo_points=True, filter_ct_domain=True):
    db_elem = _unwrap(elem)
    # If you only want CT fittings, uncomment next line to skip non-fittings:
    # if not _is_ct_fitting(db_elem): return []
    conns = _collect_connectors(db_elem, filter_ct_domain)
    pts = []
    for c in conns:
        try:
            org = c.Origin
        except Exception:
            continue
        pts.append(_as_ds_point(org) if as_dynamo_points else _xyz_tuple(org))
    return pts

def _map_deep(data, fn):
    if isinstance(data, (list, tuple)):
        return [ _map_deep(x, fn) for x in data ]
    return fn(data)

# --------------------------
# Inputs
# --------------------------
elements_in = IN[0]
as_dynamo_points = True if (len(IN) < 2 or IN[1] is None) else bool(IN[1])
filter_ct_domain = True if (len(IN) < 3 or IN[2] is None) else bool(IN[2])

# --------------------------
# Run
# --------------------------
OUT = _map_deep(elements_in, lambda e: _get_points_for_elem(e, as_dynamo_points, filter_ct_domain))