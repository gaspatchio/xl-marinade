# ABOUTME: Graph building and traversal utilities for two-pass labelling
# ABOUTME: Provides topological ordering and parent/child queries for context-aware labelling

import heapq
import logging
import sqlite3
from bisect import bisect_right
from collections import deque
from pathlib import Path

from xl_marinade.core.db_uri import connect_read_only
from xl_marinade.core.ref_converter import parse_cell_address
from xl_marinade.docs.utils.ir_schema import detect_dependency_edges

try:
    import networkx as nx

    HAS_NETWORKX = True
except ImportError:
    HAS_NETWORKX = False
    nx = None

logger = logging.getLogger(__name__)


def _fetch_range_binding_edges(conn: sqlite3.Connection) -> list[tuple[str, str]]:
    """Range-derived binding edges (Issue #2: now redundant — kept as a no-op).

    The IR producer (grouping_native._write_binding_edges_from_cells) already
    collapses every range_edge into the persisted binding_edges table (kinds
    range_static/range_dynamic), via the same populated-cell overlap this function
    recomputed at query time. Measured gained-over-persisted is 0 on the model
    DBs, and re-deriving here would add kind-less duplicates that bypass the
    provenance now carried on binding_edges. So persisted binding_edges is the
    single source of truth; this returns []. (Retained, not deleted, so callers
    and the geometric-overlap helper below stay importable.)"""
    return []

    try:  # noqa: unreachable -- legacy query-time rollup, superseded by the producer (Issue #2)
        range_edges = conn.execute("""
            SELECT from_cell_id, to_sheet_id, to_r1, to_c1, to_r2, to_c2
            FROM range_edges
        """).fetchall()
        cell_to_binding_rows = conn.execute("""
            SELECT cell_id, binding_id
            FROM cell_to_binding
        """).fetchall()
        bindings = conn.execute("""
            SELECT b.binding_id, b.sheet_id, b.address_a1, b.shape_rows, b.shape_cols, c.row, c.col
            FROM bindings b
            LEFT JOIN cells c ON c.cell_id = b.top_left_cell_id
        """).fetchall()
    except sqlite3.OperationalError:
        return []

    if not range_edges or not bindings or not cell_to_binding_rows:
        return []

    cell_to_binding: dict[int, list[str]] = {}
    for cell_id, binding_id in cell_to_binding_rows:
        cell_to_binding.setdefault(cell_id, []).append(binding_id)

    range_sheet_ids = {sheet_id for _, sheet_id, _, _, _, _ in range_edges}
    if not range_sheet_ids:
        return []

    bindings_by_sheet: dict[int, list[tuple[str, int, int, int, int]]] = {}
    for binding_id, sheet_id, address_a1, shape_rows, shape_cols, row, col in bindings:
        if sheet_id not in range_sheet_ids:
            continue
        if row is not None and col is not None and row > 0 and col > 0:
            r1 = row
            c1 = col
            r2 = row + shape_rows - 1
            c2 = col + shape_cols - 1
        else:
            parsed = parse_cell_address(address_a1)
            row = int(parsed.get("row", 0))
            col = int(parsed.get("col", 0))
            if row <= 0 or col <= 0:
                continue
            height = int(parsed.get("height", 1))
            width = int(parsed.get("width", 1))
            r1 = row
            c1 = col
            r2 = row + height - 1
            c2 = col + width - 1

        bindings_by_sheet.setdefault(sheet_id, []).append((binding_id, r1, c1, r2, c2))

    binding_index_by_sheet: dict[int, tuple[list[int], list[tuple[str, int, int, int, int]]]] = {}
    for sheet_id, targets in bindings_by_sheet.items():
        targets.sort(key=lambda item: item[1])
        binding_index_by_sheet[sheet_id] = ([item[1] for item in targets], targets)

    binding_edges: set[tuple[str, str]] = set()
    for from_cell_id, to_sheet_id, to_r1, to_c1, to_r2, to_c2 in range_edges:
        from_bindings = cell_to_binding.get(from_cell_id, [])
        if not from_bindings:
            continue
        index_entry = binding_index_by_sheet.get(to_sheet_id)
        if not index_entry:
            continue
        r1_index, targets = index_entry
        stop = bisect_right(r1_index, to_r2)
        if stop <= 0:
            continue
        for to_binding, r1, c1, r2, c2 in targets[:stop]:
            if r2 >= to_r1 and c1 <= to_c2 and c2 >= to_c1:
                for from_binding in from_bindings:
                    if from_binding != to_binding:
                        binding_edges.add((from_binding, to_binding))

    return sorted(binding_edges)


def _fetch_root_binding_ids(conn: sqlite3.Connection, cell_addr: str) -> list[str]:
    """Resolve root binding IDs for a cell address across schema variants."""
    parsed = parse_cell_address(cell_addr)
    row = int(parsed.get("row", 0))
    col = int(parsed.get("col", 0))
    height = int(parsed.get("height", 1))
    width = int(parsed.get("width", 1))
    sheet_name = parsed.get("sheet", "")
    is_range = height > 1 or width > 1

    if is_range and sheet_name and row > 0 and col > 0:
        sheet_row = conn.execute(
            """
            SELECT sheet_id
            FROM sheets
            WHERE sheet_name = ?
        """,
            (sheet_name,),
        ).fetchone()
        if sheet_row:
            sheet_id = sheet_row[0]
            r2 = row + height - 1
            c2 = col + width - 1
            rows = conn.execute(
                """
                SELECT DISTINCT ctb.binding_id
                FROM cells c
                JOIN cell_to_binding ctb ON c.cell_id = ctb.cell_id
                WHERE c.sheet_id = ?
                  AND c.row BETWEEN ? AND ?
                  AND c.col BETWEEN ? AND ?
            """,
                (sheet_id, row, r2, col, c2),
            ).fetchall()
            return [row[0] for row in rows]

    try:
        binding_rows = conn.execute(
            """
            SELECT DISTINCT ctb.binding_id
            FROM agent_cells ac
            JOIN cell_to_binding ctb ON ac.cell_id = ctb.cell_id
            WHERE ac.cell_address = ?
        """,
            (cell_addr,),
        ).fetchall()
        return [row[0] for row in binding_rows]
    except sqlite3.OperationalError:
        pass

    try:
        binding_rows = conn.execute(
            """
            SELECT DISTINCT binding_id
            FROM cells
            WHERE cell_address_a1 = ?
        """,
            (cell_addr,),
        ).fetchall()
        return [row[0] for row in binding_rows]
    except sqlite3.OperationalError:
        return []


def _fetch_binding_edges(conn: sqlite3.Connection) -> list[tuple[str, str]]:
    """Fetch binding-level edges via the canonical schema detector.

    ``detect_dependency_edges`` resolves the right edge table + columns across
    schema variants — new_arch (the current public path) persists edges in
    ``binding_edges``; the legacy schema used ``agent_binding_dependencies`` /
    ``binding_level_edges`` — using an existence check, so this works on every
    generation and stays consistent with the other ``xl_marinade.docs`` helpers.
    """
    spec = detect_dependency_edges(conn)
    if spec is None:
        return []
    return conn.execute(
        f"SELECT DISTINCT {spec.from_col}, {spec.to_col} FROM {spec.table} "
        f"ORDER BY {spec.from_col}, {spec.to_col}"
    ).fetchall()


def build_dependency_graph(ir_db_path: str) -> "nx.DiGraph | dict[str, set[str]]":
    """
    Build dependency graph from Phase 1 IR database.

    Rooted extraction: only includes bindings reachable from ``user_roots``
    (Story 15) — orphan bindings are excluded. Full-workbook extraction (the
    default public path) writes no ``user_roots``; in that case the whole
    binding-edge graph is included, so the two-pass labeller still gets its
    dependency context.

    Uses networkx if available, otherwise falls back to adjacency dict.

    Args:
        ir_db_path: Path to Phase 1 IR database

    Returns:
        NetworkX DiGraph if networkx available, else dict mapping binding_id -> set of children

    Raises:
        FileNotFoundError: If IR database doesn't exist
    """
    if not Path(ir_db_path).exists():
        raise FileNotFoundError(f"IR database not found: {ir_db_path}")

    # Read-only connection
    conn = connect_read_only(ir_db_path)

    # Get root binding IDs from user_roots
    try:
        root_cells = conn.execute("""
            SELECT sheet, range_a1 FROM user_roots
        """).fetchall()
    except sqlite3.OperationalError as e:
        if "no such table: user_roots" in str(e):
            # Backward compatibility: if user_roots doesn't exist, treat all bindings as orphans
            logger.warning(
                "user_roots table not found - treating all bindings as orphans (backward compatibility)"
            )
            conn.close()
            if HAS_NETWORKX:
                return nx.DiGraph()
            else:
                return {"__reverse__": {}}
        raise

    root_binding_ids = set()
    for sheet, range_a1 in root_cells:
        # Find binding containing this cell via cell_to_binding
        cell_addr = f"{sheet}!{range_a1}"
        binding_ids = _fetch_root_binding_ids(conn, cell_addr)
        if not binding_ids:
            logger.warning(f"Could not resolve root binding for {cell_addr}")
            continue
        root_binding_ids.update(binding_ids)

    # An empty root set is NOT an error: full-workbook extraction (the default
    # public path) writes no user_roots. In that case the whole dependency graph
    # is in scope (handled below). Only when the caller supplied roots that
    # resolved to bindings do we prune to their reachable set (Story 15).

    # Query ALL binding-level dependencies
    # from_binding depends on to_binding (from references to in its formula)
    binding_edges = _fetch_binding_edges(conn)
    range_edges = _fetch_range_binding_edges(conn)
    if range_edges:
        logger.info(f"Including {len(range_edges)} range-derived binding edges in reachability")
    all_edges = sorted(set(binding_edges) | set(range_edges))

    conn.close()

    # Build full graph first (needed for reachability computation)
    full_graph: dict[str, set[str]] = {}
    for from_id, to_id in all_edges:
        if from_id not in full_graph:
            full_graph[from_id] = set()
        full_graph[from_id].add(to_id)

    edge_bindings = set(from_id for from_id, _ in all_edges) | set(to_id for _, to_id in all_edges)

    if root_binding_ids:
        # Rooted extraction (Story 15): keep only bindings reachable from the
        # supplied roots. from → to means "from depends on to", so traverse
        # forward to collect all transitive dependencies.
        reachable = set(root_binding_ids)
        queue = deque(root_binding_ids)
        visited_count = 0
        while queue:
            current = queue.popleft()
            visited_count += 1
            if current in full_graph:
                for child in full_graph[current]:
                    if child not in reachable:
                        reachable.add(child)
                        queue.append(child)
        logger.info(
            f"Reachability BFS: {len(root_binding_ids)} roots, visited {visited_count} "
            f"nodes, found {len(reachable)} reachable bindings "
            f"(from {len(edge_bindings)} total in edges)"
        )
    else:
        # Full-workbook extraction: no user roots, so the whole dependency graph
        # is in scope. Topological ordering then starts from the natural DAG
        # roots (bindings that depend on nothing).
        reachable = edge_bindings
        logger.info(
            f"No user roots (full-workbook): including all {len(edge_bindings)} "
            f"bindings from {len(all_edges)} edges"
        )

    # Filter edges to only include reachable bindings
    reachable_edges = [
        (from_id, to_id)
        for from_id, to_id in all_edges
        if from_id in reachable and to_id in reachable
    ]

    if HAS_NETWORKX:
        # Use NetworkX DiGraph
        graph = nx.DiGraph()

        # Add edges (from -> to means "from depends on to")
        for from_id, to_id in reachable_edges:
            graph.add_edge(from_id, to_id)

        logger.info(
            f"Built NetworkX graph: {graph.number_of_nodes()} nodes, "
            f"{graph.number_of_edges()} edges"
        )
        return graph
    else:
        # Fall back to adjacency dict
        # Store both forward (dependencies) and reverse (dependents) for efficiency
        graph: dict[str, set[str]] = {}
        reverse_graph: dict[str, set[str]] = {}

        for from_id, to_id in reachable_edges:
            # from_id depends on to_id
            if from_id not in graph:
                graph[from_id] = set()
            graph[from_id].add(to_id)

            # to_id is depended on by from_id
            if to_id not in reverse_graph:
                reverse_graph[to_id] = set()
            reverse_graph[to_id].add(from_id)

        # Store reverse graph in special key for later use
        graph["__reverse__"] = reverse_graph

        logger.info(f"Built adjacency dict graph: {len(graph) - 1} nodes")
        return graph


def _kahn_order(graph: dict[str, set[str]]) -> list[str]:
    """Kahn's algorithm over the dict graph shape, cycle-tolerant.

    Nodes on (or downstream of) a cycle never reach in-degree 0; they are
    appended at the end in sorted order so every node appears exactly once.
    """
    # Build in-degree map
    reverse_graph: dict[str, set[str]] = graph.get("__reverse__", {})
    all_nodes = set(graph.keys()) - {"__reverse__"}
    all_nodes.update(reverse_graph.keys())

    in_degree = dict.fromkeys(all_nodes, 0)

    # Count in-degrees
    for node in all_nodes:
        if node in graph and node != "__reverse__":
            for child in graph[node]:
                if child in in_degree:
                    in_degree[child] += 1

    # Start with nodes that have no dependencies (in_degree = 0)
    queue = [node for node, degree in in_degree.items() if degree == 0]
    heapq.heapify(queue)
    result = []

    while queue:
        # Pop from min-heap for determinism
        node = heapq.heappop(queue)
        result.append(node)

        # Reduce in-degree of children
        if node in graph and node != "__reverse__":
            children = sorted(graph[node])  # Sort for determinism
            for child in children:
                if child in in_degree:
                    in_degree[child] -= 1
                    if in_degree[child] == 0:
                        heapq.heappush(queue, child)

    # If result is shorter than all_nodes, there are cycles
    if len(result) < len(all_nodes):
        # Add remaining nodes in sorted order
        remaining = sorted(set(all_nodes) - set(result))
        logger.warning(f"Graph contains cycles, {len(remaining)} nodes have circular dependencies")
        result.extend(remaining)

    return result


def get_topological_order(graph: "nx.DiGraph | dict[str, set[str]]") -> list[str]:
    """
    Get topological ordering of bindings (roots to leaves).

    Binding-level graphs are legitimately cyclic for recursive models: a
    time-lagged projection column is a DAG cell-by-cell but a cycle once
    collapsed to column bindings. For cyclic graphs both backends return the
    same ordering: the acyclic portion in dependency order, then the nodes
    stuck on (or downstream of) cycles appended in sorted order.

    Args:
        graph: Dependency graph from build_dependency_graph

    Returns:
        List of binding IDs in topological order (roots first, leaves last)
    """
    if HAS_NETWORKX and isinstance(graph, nx.DiGraph):
        try:
            return list(nx.topological_sort(graph))
        except nx.NetworkXUnfeasible:
            # Cyclic — fall back to the dict backend's cycle-tolerant Kahn's
            # ordering so both backends stay identical on the same graph.
            logger.warning("Graph contains cycles, using Kahn's ordering with cyclic remainder")
            dict_graph: dict[str, set[str]] = {
                node: set(graph.successors(node)) for node in graph.nodes()
            }
            dict_graph["__reverse__"] = {
                node: set(graph.predecessors(node)) for node in graph.nodes()
            }
            return _kahn_order(dict_graph)
    else:
        return _kahn_order(graph)


def get_parents(graph: "nx.DiGraph | dict[str, set[str]]", binding_id: str) -> list[str]:
    """
    Get parent bindings (bindings that depend on this binding).

    Parent = depends on me = references me in formula

    Args:
        graph: Dependency graph from build_dependency_graph
        binding_id: Binding ID to query

    Returns:
        List of parent binding IDs (sorted for determinism)
    """
    if HAS_NETWORKX and isinstance(graph, nx.DiGraph):
        # In NetworkX, predecessors are nodes that have edges TO this node
        # Edge from->to means "from depends on to"
        # So predecessors of 'to' are the parents (things that depend on it)
        if binding_id in graph:
            return sorted(graph.predecessors(binding_id))
        else:
            return []
    else:
        # Use reverse graph stored during construction
        reverse_graph: dict[str, set[str]] = graph.get("__reverse__", {})
        if binding_id in reverse_graph:
            return sorted(reverse_graph[binding_id])
        else:
            return []


def get_children(graph: "nx.DiGraph | dict[str, set[str]]", binding_id: str) -> list[str]:
    """
    Get child bindings (bindings that this binding depends on).

    Child = I depend on = I reference in my formula

    Args:
        graph: Dependency graph from build_dependency_graph
        binding_id: Binding ID to query

    Returns:
        List of child binding IDs (sorted for determinism)
    """
    if HAS_NETWORKX and isinstance(graph, nx.DiGraph):
        # In NetworkX, successors are nodes that have edges FROM this node
        # Edge from->to means "from depends on to"
        # So successors of 'from' are the children (things it depends on)
        if binding_id in graph:
            return sorted(graph.successors(binding_id))
        else:
            return []
    else:
        # Use forward graph
        if binding_id in graph and binding_id != "__reverse__":
            return sorted(graph[binding_id])
        else:
            return []
