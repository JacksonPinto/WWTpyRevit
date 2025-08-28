#! python3
# calc_shortest.py
# Topologic-based shortest path calculator (restored use of topologicpy)
# Version: 4.0.0 (2025-08-29)
# Author: JacksonPinto
#
# GOAL:
#   Use the official topologicpy API (Graph.ByVerticesEdges + Graph.ShortestPath)
#   to compute shortest (metric) paths from every start point to the single end point
#   exactly over the edges listed in topologic.JSON (NO topology rebuilding /
#   clustering that can delete or rewire edges).
#
# KEY POINTS / DIFFERENCES FROM 3.x Pure-Python:
#   - Builds Graph directly with Graph.ByVerticesEdges(vertices, edges).
#   - Uses edgeKey='Length' so geometric 3D length weights are applied.
#   - Avoids Cluster/SelfMerge (those were collapsing/merging vertices/edges).
#   - Optional bridging of start/end points to nearest vertex if they are not
#     already part of the graph (configurable).
#   - Produces results compatible with your Dynamo reader (element_id, length,
#     vertex_path_xyz, plus status flags).
#
# INPUT EXPECTED (topologic.JSON):
# {
#   "vertices":[ [x,y,z], ... ],
#   "edges":[ [i,j], ... ],
#   "start_points":[ {"element_id":123, "point":[x,y,z]}, ... ],
#   "end_point":[x,y,z],
#   "meta": {... optional ...}
# }
#
# OUTPUT (topologic_results.json):
# {
#   "meta":{ ... diagnostics ... },
#   "end_point":[x,y,z],
#   "results":[
#     {
#       "start_index": i,
#       "element_id": <id>,
#       "length": <float or null>,
#       "vertex_indices": [v0,v1,...],          # indices referencing ORIGINAL vertex list (best effort)
#       "vertex_path_xyz":[[x,y,z],...],
#       "status":"ok" | "no_path" | "unmatched_start" | "bridged_start" | "bridged_end"
#     }, ...
#   ]
# }
#
# CONFIGURATION:
#   TOLERANCE_GLOBAL        : passed to topologic core (very small to reduce unintended merges)
#   MATCH_TOL               : tolerance for matching start/end coords to existing vertices
#   ALLOW_BRIDGING_STARTS   : if a start point coordinate is not an exact match, create a new vertex
#                             and an edge to nearest existing vertex (straight segment).
#   ALLOW_BRIDGING_END      : same logic for end point if not found.
#   BRIDGING_MAX_DIST       : maximum allowed distance (feet) to connect bridging; if exceeded -> no_path.
#   IGNORE_DUPLICATE_EDGES  : skip adding reverse duplicates (unordered).
#
# NOTE:
#   If you still see "missing" infrastructure in the results visualization remember:
#   - Only SHORTEST PATH edges per start are exported, not entire infrastructure.
#   To visualize entire network load topologic.JSON (blue) separately from path results (red).
#
# DEPENDENCIES:
#   pip install topologicpy  (already in your environment)
#
# ------------------------------------------------------------------------------

import os, sys, json, math, time

# ---------------- CONFIG ----------------
TOLERANCE_GLOBAL        = 1e-9
MATCH_TOL               = 1e-9
ALLOW_BRIDGING_STARTS   = True
ALLOW_BRIDGING_END      = True
BRIDGING_MAX_DIST       = 5.0      # feet (set higher if devices can be far from nearest vertex)
IGNORE_DUPLICATE_EDGES  = True
EDGE_KEY                = "Length" # used in Graph.ShortestPath
# ----------------------------------------

# ---------------- IMPORT TOPOLOGICPY ----------------
try:
    from topologicpy.Vertex import Vertex
    from topologicpy.Edge import Edge
    from topologicpy.Graph import Graph
    import topologic  # core
except Exception as e:
    print("[ERROR] topologicpy import failed: {}".format(e))
    sys.exit(1)

# Set global tolerance (try/except in case API changes)
try:
    topologic.Topology.Tolerance(TOLERANCE_GLOBAL)
    print("[DEBUG] Set topologic core tolerance to {}".format(TOLERANCE_GLOBAL))
except Exception as e:
    print("[WARN] Could not set topologic tolerance: {}".format(e))

# ---------------- UTILS ----------------
def dist3(a,b):
    return math.sqrt((a[0]-b[0])**2+(a[1]-b[1])**2+(a[2]-b[2])**2)

def load_json(path):
    with open(path,'r') as f:
        return json.load(f)

def save_json(path, data):
    with open(path,'w') as f:
        json.dump(data, f, indent=2)

def coord_match_index(coord, vertices_data, tol):
    # Return first index within tol
    cx,cy,cz = coord
    for i,(x,y,z) in enumerate(vertices_data):
        if abs(x-cx)<=tol and abs(y-cy)<=tol and abs(z-cz)<=tol:
            return i
    return None

def nearest_vertex_index(coord, vertices_data):
    best_i=None; best_d=None
    for i,(x,y,z) in enumerate(vertices_data):
        d = ( (x-coord[0])**2 + (y-coord[1])**2 + (z-coord[2])**2 )**0.5
        if best_d is None or d<best_d:
            best_d=d; best_i=i
    return best_i, best_d

# Mapping from Vertex objects back to original indices:
def vertex_obj_index_map(original_vertices_data, created_vertex_objs, tol):
    mapping={}
    for idx, vdata in enumerate(original_vertices_data):
        for vobj in created_vertex_objs:
            try:
                x,y,z = vobj.Coordinates()
                if abs(x-vdata[0])<=tol and abs(y-vdata[1])<=tol and abs(z-vdata[2])<=tol:
                    mapping[id(vobj)] = idx
                    break
            except:
                pass
    return mapping

# ---------------- MAIN ----------------
def main():
    t0 = time.time()
    script_dir = os.path.dirname(os.path.abspath(__file__))
    in_path  = os.path.join(script_dir, "topologic.JSON")
    out_path = os.path.join(script_dir, "topologic_results.json")

    if not os.path.exists(in_path):
        print("[ERROR] Input JSON not found: {}".format(in_path))
        sys.exit(1)

    data = load_json(in_path)
    vertices_data = data.get("vertices", [])
    edges_data    = data.get("edges", [])
    start_points  = data.get("start_points", [])
    end_point     = data.get("end_point")
    input_meta    = data.get("meta", {})

    if not vertices_data or not edges_data:
        print("[ERROR] vertices or edges list empty.")
        # We can still attempt direct device->end length, but that is not requested.
        sys.exit(1)
    if end_point is None:
        print("[ERROR] end_point missing.")
        sys.exit(1)

    print("[INFO] Building vertices: {}".format(len(vertices_data)))
    vertex_objs = []
    for v in vertices_data:
        try:
            vertex_objs.append(Vertex.ByCoordinates(v[0], v[1], v[2]))
        except Exception as e:
            print("[WARN] Vertex creation failed for {}: {}".format(v, e))
            vertex_objs.append(None)

    # Build edges, optionally skip duplicates
    seen_pairs=set()
    edge_objs=[]
    valid_edge_count=0
    skipped_duplicate=0
    skipped_invalid=0
    for idx,(i,j) in enumerate(edges_data):
        try:
            i2=int(i); j2=int(j)
        except:
            skipped_invalid+=1
            continue
        if i2<0 or j2<0 or i2>=len(vertex_objs) or j2>=len(vertex_objs):
            skipped_invalid+=1
            continue
        if vertex_objs[i2] is None or vertex_objs[j2] is None:
            skipped_invalid+=1
            continue
        if i2==j2:
            skipped_invalid+=1
            continue
        key = (i2,j2) if i2<j2 else (j2,i2)
        if IGNORE_DUPLICATE_EDGES and key in seen_pairs:
            skipped_duplicate+=1
            continue
        seen_pairs.add(key)
        try:
            e = Edge.ByVertices(vertex_objs[i2], vertex_objs[j2])
            edge_objs.append(e)
            valid_edge_count+=1
        except Exception as e:
            print("[WARN] Edge creation failed idx {} ({}-{}): {}".format(idx,i2,j2,e))
            skipped_invalid+=1

    print("[INFO] Edges built: used={} dupSkipped={} invalidSkipped={}".format(
        valid_edge_count, skipped_duplicate, skipped_invalid))

    # Build Graph without merging (Graph.ByVerticesEdges)
    try:
        graph = Graph.ByVerticesEdges(vertex_objs, edge_objs)
    except Exception as e:
        print("[ERROR] Graph.ByVerticesEdges failed: {}".format(e))
        sys.exit(1)

    # Build a mapping (id(vertex_obj) -> original index) for path index reconstruction
    vobj_to_index = vertex_obj_index_map(vertices_data, vertex_objs, MATCH_TOL)

    # Find end vertex
    end_vid = coord_match_index(end_point, vertices_data, MATCH_TOL)
    end_status = "exact"
    if end_vid is None:
        if ALLOW_BRIDGING_END:
            ni, nd = nearest_vertex_index(end_point, vertices_data)
            if nd is not None and nd <= BRIDGING_MAX_DIST:
                # Bridge by adding a new vertex and edge (straight).
                print("[INFO] Bridging end point (dist {:.4f} ft)".format(nd))
                try:
                    end_vertex_obj = Vertex.ByCoordinates(end_point[0], end_point[1], end_point[2])
                    nearest_existing = vertex_objs[ni]
                    bridge_edge = Edge.ByVertices(nearest_existing, end_vertex_obj)
                    # Add vertex then edge
                    graph = Graph.AddVertex(graph, end_vertex_obj)
                    graph = Graph.AddEdge(graph, bridge_edge)
                    vertex_objs.append(end_vertex_obj)
                    vertices_data.append(end_point[:])
                    end_vid = len(vertices_data)-1
                    end_status = "bridged_end"
                    vobj_to_index[id(end_vertex_obj)] = end_vid
                except Exception as e:
                    print("[ERROR] Failed bridging end point: {}".format(e))
                    sys.exit(1)
            else:
                print("[ERROR] End point not found and not within bridging distance.")
                sys.exit(1)
        else:
            print("[ERROR] End point not found (exact) and bridging disabled.")
            sys.exit(1)

    end_vertex_obj = vertex_objs[end_vid]

    results=[]
    success=0
    fail=0
    bridged_starts=0
    unmatched_starts=0
    max_len=0.0
    min_len=None
    sum_len=0.0
    path_lengths=[]

    for si, sp in enumerate(start_points):
        eid = sp.get("element_id")
        coord = sp.get("point")
        status="ok"
        path_vertex_indices=[]
        length_value=None

        if not coord:
            status="unmatched_start"
            unmatched_starts+=1
            fail+=1
            results.append({
                "start_index": si,
                "element_id": eid,
                "length": None,
                "vertex_indices": [],
                "vertex_path_xyz": [],
                "status": status
            })
            continue

        # Match start
        s_vid = coord_match_index(coord, vertices_data, MATCH_TOL)
        if s_vid is None:
            # Attempt bridging if allowed
            if ALLOW_BRIDGING_STARTS:
                ni, nd = nearest_vertex_index(coord, vertices_data)
                if nd is not None and nd <= BRIDGING_MAX_DIST:
                    try:
                        new_v = Vertex.ByCoordinates(coord[0], coord[1], coord[2])
                        nearest_existing = vertex_objs[ni]
                        bridge_edge = Edge.ByVertices(nearest_existing, new_v)
                        graph = Graph.AddVertex(graph, new_v)
                        graph = Graph.AddEdge(graph, bridge_edge)
                        vertex_objs.append(new_v)
                        vertices_data.append(coord[:])
                        s_vid = len(vertices_data)-1
                        vobj_to_index[id(new_v)] = s_vid
                        status="bridged_start"
                        bridged_starts+=1
                    except Exception as e:
                        print("[WARN] Failed to bridge start {}: {}".format(eid,e))
                        status="unmatched_start"
                        unmatched_starts+=1
                else:
                    status="unmatched_start"
                    unmatched_starts+=1
            else:
                status="unmatched_start"
                unmatched_starts+=1

        if s_vid is None:
            fail+=1
            results.append({
                "start_index": si,
                "element_id": eid,
                "length": None,
                "vertex_indices": [],
                "vertex_path_xyz": [],
                "status": status
            })
            continue

        start_vertex_obj = vertex_objs[s_vid]

        # Compute shortest path via Topologic
        try:
            wire = Graph.ShortestPath(graph, start_vertex_obj, end_vertex_obj, edgeKey=EDGE_KEY)
        except Exception as e:
            print("[ERROR] ShortestPath error element {}: {}".format(eid,e))
            wire = None

        if not wire:
            status="no_path"
            fail+=1
            results.append({
                "start_index": si,
                "element_id": eid,
                "length": None,
                "vertex_indices": [],
                "vertex_path_xyz": [],
                "status": status
            })
            continue

        # Extract path vertices & compute metric length (sum of straight segments)
        try:
            from topologicpy.Wire import Wire as TPWire
            wverts = TPWire.Vertices(wire)
        except Exception:
            # Older import path fallback
            from topologicpy.Wire import Wire
            wverts = Wire.Vertices(wire)

        coords=[]
        for v in wverts:
            try:
                coords.append(list(v.Coordinates()))
            except:
                coords.append([None,None,None])

        # compute length
        pl=0.0
        for i in range(len(coords)-1):
            a=coords[i]; b=coords[i+1]
            if None in a or None in b: continue
            pl+=dist3(a,b)

        length_value=pl
        path_lengths.append(pl)
        sum_len+=pl
        if min_len is None or pl<min_len: min_len=pl
        if pl>max_len: max_len=pl
        success+=1

        # Map vertex objects to original indices (best effort via coordinate match)
        # We'll loop each coord to find first matching index; O(N*pathLen) but path lengths are small.
        for c in coords:
            vidx = coord_match_index(c, vertices_data, MATCH_TOL)
            if vidx is None:
                # Might be a bridged vertex appended after mapping creation
                # Already added to vertices_data with exact coords, so second try:
                vidx = coord_match_index(c, vertices_data, 1e-7)
            path_vertex_indices.append(vidx if vidx is not None else -1)

        results.append({
            "start_index": si,
            "element_id": eid,
            "length": length_value,
            "vertex_indices": path_vertex_indices,
            "vertex_path_xyz": coords,
            "status": status
        })

    avg_len = (sum_len / success) if success>0 else None
    meta_out = {
        "input_meta": input_meta,
        "vertex_count_initial": len(data.get("vertices", [])),
        "vertex_count_after_bridging": len(vertices_data),
        "edge_count_input": len(edges_data),
        "edge_count_used": valid_edge_count,
        "duplicate_edges_skipped": skipped_duplicate,
        "invalid_edges_skipped": skipped_invalid,
        "start_count": len(start_points),
        "end_status": end_status,
        "paths_success": success,
        "paths_failed": fail,
        "bridged_starts": bridged_starts,
        "unmatched_starts": unmatched_starts,
        "min_path_length": min_len,
        "max_path_length": max_len if success>0 else None,
        "avg_path_length": avg_len,
        "edge_key_used": EDGE_KEY,
        "tolerance_global": TOLERANCE_GLOBAL,
        "match_tol": MATCH_TOL,
        "bridging_enabled_starts": ALLOW_BRIDGING_STARTS,
        "bridging_enabled_end": ALLOW_BRIDGING_END,
        "bridging_max_dist": BRIDGING_MAX_DIST,
        "runtime_seconds": round(time.time()-t0,4),
        "notes": "Graph built with Graph.ByVerticesEdges; no SelfMerge/Cluster. Paths use Graph.ShortestPath with edgeKey='Length'."
    }

    payload = {
        "meta": meta_out,
        "end_point": end_point,
        "results_count": len(results),
        "paths_success": success,
        "paths_failed": fail,
        "results": results
    }

    save_json(out_path, payload)
    print("[INFO] Shortest path computation complete.")
    print("       Success:{}  Failed:{}  BridgedStarts:{}  UnmatchedStarts:{}"
          .format(success, fail, bridged_starts, unmatched_starts))
    print("       Output: {}".format(out_path))

if __name__ == "__main__":
    main()