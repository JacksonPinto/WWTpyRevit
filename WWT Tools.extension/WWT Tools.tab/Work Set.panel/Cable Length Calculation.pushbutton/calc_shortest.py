#! python3
# calc_shortest.py
# Version: 3.2.1
# Author: JacksonPinto
# Description:
#   Reads topologic.JSON (vertices/edges format), builds a fully TopologicPy graph object,
#   maps start points and end point to Topologic graph vertices,
#   and computes shortest paths using TopologicPy's methods.
# Update 3.2.1:
#   - Sets Topologic global tolerance to 0.001 for geometric operations

import json
import sys
import os
import math

from topologicpy.Topology import Topology
from topologicpy.Vertex import Vertex
from topologicpy.Edge import Edge
from topologicpy.Cluster import Cluster
from topologicpy.Graph import Graph
from topologicpy.Wire import Wire

# Set Topologic tolerance BEFORE any geometric operation
try:
    import topologic
    topologic.Topology.Tolerance(0.001)
    print("[DEBUG] Topologic global tolerance set to 0.001")
except Exception as e:
    print(f"[DEBUG] Could not set Topologic tolerance: {e}")

def dist3(a, b):
    """Euclidean distance between two 3D points."""
    return math.sqrt((a[0]-b[0])**2 + (a[1]-b[1])**2 + (a[2]-b[2])**2)

def nearest_vertex_idx(pt, coords_list):
    """Return index of point in coords_list nearest to pt."""
    return min(range(len(coords_list)), key=lambda i: dist3(pt, coords_list[i]))

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    json_path = os.path.join(script_dir, "topologic.JSON")
    results_path = os.path.join(script_dir, "topologic_results.json")
    print("[DEBUG] Script started")
    print(f"[DEBUG] Loading JSON from: {json_path}")

    # Load input JSON
    with open(json_path, 'r') as f:
        data = json.load(f)
    vertices_json = data.get("vertices", [])
    edges_json = data.get("edges", [])
    start_points = data.get("start_points", [])
    end_point = data.get("end_point", None)
    print(f"[DEBUG] Input vertices: {len(vertices_json)}")
    print(f"[DEBUG] Input edges: {len(edges_json)}")
    print(f"[DEBUG] Start points: {len(start_points)}")
    print(f"[DEBUG] End point: {end_point}")

    # 1. Create Topologic Vertex and Edge objects
    vertices = [Vertex.ByCoordinates(*v) for v in vertices_json]
    edges = []
    for i, (a, b) in enumerate(edges_json):
        try:
            edges.append(Edge.ByVertices(vertices[a], vertices[b]))
        except Exception as e:
            print(f"[DEBUG] Error creating edge {i}: {e}")

    print(f"[DEBUG] Created {len(vertices)} Topologic vertices, {len(edges)} Topologic edges.")

    # 2. Create a Cluster from Edges
    cluster = Cluster.ByTopologies(*edges)
    print("[DEBUG] Cluster created.")

    # 3. Self-merge the Cluster to a logical topology
    logical_topology = Topology.SelfMerge(cluster)
    print("[DEBUG] Self-merged topology created.")

    # 4. Create a Graph from the logical topology
    graph = Graph.ByTopology(logical_topology)
    print("[DEBUG] Topologic graph object created.")

    # 5. Get the list of graph vertices and their coordinates
    graph_vertices = Graph.Vertices(graph)
    graph_coords = [v.Coordinates() for v in graph_vertices]
    print(f"[DEBUG] Graph vertex count: {len(graph_vertices)}")

    # 6. Map start/end points to nearest graph vertex
    if not graph_vertices:
        print("[ERROR] No vertices in Topologic graph! Exiting.")
        sys.exit(1)

    # Map end point
    if end_point is not None:
        end_idx = nearest_vertex_idx(end_point, graph_coords)
        end_vertex = graph_vertices[end_idx]
        print(f"[DEBUG] End point maps to graph vertex index: {end_idx} coords: {graph_coords[end_idx]}")
    else:
        print("[ERROR] No end_point defined in JSON. Exiting.")
        sys.exit(1)

    # Map all start points (FIX: use sp['point'])
    start_indices = []
    start_graph_vertices = []
    for i, sp in enumerate(start_points):
        idx = nearest_vertex_idx(sp["point"], graph_coords)
        start_indices.append(idx)
        start_graph_vertices.append(graph_vertices[idx])
        print("[DEBUG] Start point {} maps to graph vertex index: {} coords: {}".format(i, idx, graph_coords[idx]))

    # 7. Compute shortest path from each start point to end point
    results = []
    for i, v_start in enumerate(start_graph_vertices):
        try:
            # Returns a 'Wire' topology (path), or None if no path found
            path_wire = Graph.ShortestPath(graph, v_start, end_vertex)
            if path_wire:
                path_vertices = Wire.Vertices(path_wire)
                path_xyz = [v.Coordinates() for v in path_vertices]
                length = sum(dist3(path_xyz[j], path_xyz[j+1]) for j in range(len(path_xyz)-1))
                print(f"[DEBUG] Shortest path for start {i}: length={length:.3f}, vertex_count={len(path_xyz)}")
                results.append({
                    "start_index": i,
                    "element_id": start_points[i]["element_id"],
                    "length": length,
                    "vertex_path_xyz": path_xyz
                })
            else:
                print(f"[DEBUG] No path found for start {i}.")
                results.append({
                    "start_index": i,
                    "element_id": start_points[i]["element_id"],
                    "length": None,
                    "vertex_path_xyz": []
                })
        except Exception as e:
            print(f"[DEBUG] Error computing shortest path for start {i}: {e}")
            results.append({
                "start_index": i,
                "element_id": start_points[i]["element_id"],
                "length": None,
                "vertex_path_xyz": []
            })

    # 8. Write results
    with open(results_path, "w") as f:
        json.dump(results, f, indent=2)
    print(f"[DEBUG] Results written to: {results_path}")

if __name__ == "__main__":
    main()