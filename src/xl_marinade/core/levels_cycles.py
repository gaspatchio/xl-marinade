# ABOUTME: Compute topological levels and detect cycles in binding dependency graph.
# ABOUTME: Uses Kahn's algorithm for levels and Tarjan's algorithm for SCC detection.

from collections import defaultdict
from dataclasses import dataclass


@dataclass
class LevelAssignment:
    """
    Topological level assignment for a binding.

    Attributes:
        level: Topological level (0-indexed, 0 = no dependencies)
        binding_id: Binding ID
    """

    level: int
    binding_id: str


@dataclass
class Cycle:
    """
    Circular dependency group (strongly connected component).

    Attributes:
        cycle_id: Cycle identifier (0-indexed)
        bindings: List of binding IDs in cycle (deterministically ordered)
    """

    cycle_id: int
    bindings: list[str]


def compute_levels(
    binding_ids: list[str],
    edges: list[tuple[str, str]],
    binding_metadata: dict[str, dict[str, str | int]] | None = None,
) -> list[LevelAssignment]:
    """
    Compute topological levels using Kahn's algorithm.

    Kahn's algorithm:
    1. Start with nodes that have no incoming edges (level 0)
    2. Process nodes in deterministic order (sorted by binding metadata)
    3. Remove processed nodes and their outgoing edges
    4. Repeat until all nodes processed or cycle detected

    Args:
        binding_ids: List of all binding IDs in the graph
        edges: List of (from_binding_id, to_binding_id) tuples
        binding_metadata: Optional dict mapping binding_id to metadata dict
                         with 'sheet', 'row', 'col' for deterministic ordering

    Returns:
        List of LevelAssignment objects sorted by level, then binding_id

    Raises:
        ValueError: If graph contains cycles (use detect_cycles first)

    Determinism:
        Within each level, bindings are processed in sorted order by:
        (sheet ASC, row ASC, col ASC) if binding_metadata provided,
        otherwise by binding_id lexicographically.

    Example:
        >>> # DAG: A → B → C
        >>> binding_ids = ["A", "B", "C"]
        >>> edges = [("B", "A"), ("C", "B")]
        >>> levels = compute_levels(binding_ids, edges)
        >>> [(l.level, l.binding_id) for l in levels]
        [(0, 'A'), (1, 'B'), (2, 'C')]
    """
    # Build adjacency lists
    in_degree: dict[str, int] = dict.fromkeys(binding_ids, 0)
    out_edges: dict[str, list[str]] = defaultdict(list)

    for from_id, to_id in edges:
        # Edge from A to B means A depends on B
        # So B has an outgoing edge to A for topological ordering
        out_edges[to_id].append(from_id)
        in_degree[from_id] = in_degree.get(from_id, 0) + 1

    # Find all nodes with no dependencies (in_degree = 0)
    current_level_nodes = [bid for bid in binding_ids if in_degree[bid] == 0]

    # Sort for deterministic ordering
    current_level_nodes.sort(key=lambda bid: _sort_key(bid, binding_metadata))

    levels: list[LevelAssignment] = []
    level_num = 0
    processed_count = 0

    while current_level_nodes:
        # Assign current level
        for binding_id in current_level_nodes:
            levels.append(LevelAssignment(level=level_num, binding_id=binding_id))
            processed_count += 1

        # Find next level nodes
        next_level_nodes = []
        for binding_id in current_level_nodes:
            # Remove this node and its outgoing edges
            for dependent in out_edges[binding_id]:
                in_degree[dependent] -= 1
                if in_degree[dependent] == 0:
                    next_level_nodes.append(dependent)

        # Sort next level for determinism
        next_level_nodes.sort(key=lambda bid: _sort_key(bid, binding_metadata))

        current_level_nodes = next_level_nodes
        level_num += 1

    # Check if all nodes were processed (no cycles)
    if processed_count != len(binding_ids):
        unprocessed = [bid for bid in binding_ids if in_degree[bid] > 0]
        raise ValueError(
            f"Graph contains cycles. {len(unprocessed)} nodes not processed: "
            f"{', '.join(sorted(unprocessed)[:5])}{'...' if len(unprocessed) > 5 else ''}"
        )

    return levels


def detect_cycles(
    binding_ids: list[str],
    edges: list[tuple[str, str]],
    binding_metadata: dict[str, dict[str, str | int]] | None = None,
) -> list[Cycle]:
    """
    Detect cycles using Tarjan's strongly connected components algorithm.

    A strongly connected component (SCC) is a maximal set of nodes where
    every node is reachable from every other node. In a dependency graph,
    SCCs with size > 1 represent circular dependencies. Self-loops (A → A)
    are also detected as cycles.

    Args:
        binding_ids: List of all binding IDs in the graph
        edges: List of (from_binding_id, to_binding_id) tuples
        binding_metadata: Optional dict for deterministic ordering within cycles

    Returns:
        List of Cycle objects (SCCs with size > 1 or self-loops).
        Cycles are sorted by (size DESC, first binding_id ASC).
        Within each cycle, bindings are sorted deterministically.

    Determinism:
        - SCCs are found in deterministic order via DFS from sorted nodes
        - Within each SCC, bindings are sorted by metadata or binding_id
        - Cycle IDs are assigned in order of discovery

    Example:
        >>> # Cycle: A → B → C → A
        >>> binding_ids = ["A", "B", "C"]
        >>> edges = [("A", "B"), ("B", "C"), ("C", "A")]
        >>> cycles = detect_cycles(binding_ids, edges)
        >>> len(cycles)
        1
        >>> sorted(cycles[0].bindings)
        ['A', 'B', 'C']
    """
    # Build adjacency list (forward direction: from → to)
    adj_list: dict[str, list[str]] = defaultdict(list)
    has_self_loop: set[str] = set()

    for from_id, to_id in edges:
        if from_id == to_id:
            # Self-loop
            has_self_loop.add(from_id)
        else:
            adj_list[from_id].append(to_id)

    # Sort adjacency lists for deterministic DFS traversal
    for binding_id in adj_list:
        adj_list[binding_id].sort(key=lambda bid: _sort_key(bid, binding_metadata))

    # Tarjan's algorithm state
    index_counter = [0]
    stack: list[str] = []
    lowlinks: dict[str, int] = {}
    index: dict[str, int] = {}
    on_stack: set[str] = set()
    sccs: list[list[str]] = []

    def strongconnect(node: str) -> None:
        """Tarjan's recursive DFS helper."""
        # Set the depth index for this node
        index[node] = index_counter[0]
        lowlinks[node] = index_counter[0]
        index_counter[0] += 1
        stack.append(node)
        on_stack.add(node)

        # Consider successors
        for successor in adj_list[node]:
            if successor not in index:
                # Successor has not yet been visited; recurse
                strongconnect(successor)
                lowlinks[node] = min(lowlinks[node], lowlinks[successor])
            elif successor in on_stack:
                # Successor is in stack and hence in current SCC
                lowlinks[node] = min(lowlinks[node], index[successor])

        # If node is a root node, pop the stack and create an SCC
        if lowlinks[node] == index[node]:
            scc = []
            while True:
                successor = stack.pop()
                on_stack.remove(successor)
                scc.append(successor)
                if successor == node:
                    break
            sccs.append(scc)

    # Run Tarjan's algorithm from all unvisited nodes (sorted for determinism)
    sorted_binding_ids = sorted(binding_ids, key=lambda bid: _sort_key(bid, binding_metadata))
    for binding_id in sorted_binding_ids:
        if binding_id not in index:
            strongconnect(binding_id)

    # Filter to only cycles (SCC size > 1) and create Cycle objects
    cycles: list[Cycle] = []
    cycle_id = 0

    for scc in sccs:
        if len(scc) > 1:
            # Sort bindings within cycle deterministically
            scc_sorted = sorted(scc, key=lambda bid: _sort_key(bid, binding_metadata))
            cycles.append(Cycle(cycle_id=cycle_id, bindings=scc_sorted))
            cycle_id += 1

    # Add self-loops as single-node cycles
    for binding_id in sorted(has_self_loop, key=lambda bid: _sort_key(bid, binding_metadata)):
        cycles.append(Cycle(cycle_id=cycle_id, bindings=[binding_id]))
        cycle_id += 1

    # Sort cycles by (size DESC, first binding_id ASC) for deterministic output
    cycles.sort(key=lambda c: (-len(c.bindings), c.bindings[0]))

    # Reassign cycle IDs after sorting
    for i, cycle in enumerate(cycles):
        cycle.cycle_id = i

    return cycles


def _sort_key(
    binding_id: str, binding_metadata: dict[str, dict[str, str | int]] | None
) -> tuple[str, int, int, str] | tuple[str]:
    """
    Generate sort key for binding deterministic ordering.

    Args:
        binding_id: Binding ID
        binding_metadata: Optional metadata dict with 'sheet', 'row', 'col'

    Returns:
        Sort key tuple: (sheet, row, col) if metadata available,
                       otherwise (binding_id,)
    """
    if binding_metadata and binding_id in binding_metadata:
        meta = binding_metadata[binding_id]
        sheet = str(meta.get("sheet", ""))
        row = int(meta.get("row", 0))
        col = int(meta.get("col", 0))
        return (sheet, row, col, binding_id)
    return (binding_id,)


def levels_to_db_rows(levels: list[LevelAssignment]) -> list[tuple[int, str]]:
    """
    Convert LevelAssignment list to database rows.

    Args:
        levels: List of LevelAssignment objects

    Returns:
        List of (level, binding_id) tuples sorted by level, then binding_id

    Example:
        >>> levels = [LevelAssignment(0, "A"), LevelAssignment(1, "B")]
        >>> levels_to_db_rows(levels)
        [(0, 'A'), (1, 'B')]
    """
    rows = [(level.level, level.binding_id) for level in levels]
    rows.sort(key=lambda x: (x[0], x[1]))  # Sort by level, then binding_id
    return rows


def cycles_to_db_rows(cycles: list[Cycle]) -> list[tuple[int, int, str]]:
    """
    Convert Cycle list to database rows.

    Args:
        cycles: List of Cycle objects

    Returns:
        List of (cycle_id, ord, binding_id) tuples.
        'ord' is the position within the cycle (0-indexed).
        Sorted by cycle_id, then ord.

    Example:
        >>> cycle = Cycle(cycle_id=0, bindings=["A", "B", "C"])
        >>> cycles_to_db_rows([cycle])
        [(0, 0, 'A'), (0, 1, 'B'), (0, 2, 'C')]
    """
    rows = []
    for cycle in cycles:
        for ord_num, binding_id in enumerate(cycle.bindings):
            rows.append((cycle.cycle_id, ord_num, binding_id))

    # Already sorted by construction (cycles sorted, bindings within cycles sorted)
    return rows
