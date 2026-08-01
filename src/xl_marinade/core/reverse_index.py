# ABOUTME: Build reverse dependency index mapping precedents to their dependents
# ABOUTME: Maintains deterministic ordering for reproducible graph construction

from collections import defaultdict


class ReverseIndex:
    """
    Reverse dependency index: precedent → set of dependents.

    For each cell that is referenced (precedent), tracks all cells
    that reference it (dependents).

    Example:
        If cell B1 contains formula =A1+A2, then:
        - A1 → {B1}
        - A2 → {B1}

    This index is used for downstream traversal (Story 3).
    """

    def __init__(self) -> None:
        """Initialize empty reverse index"""
        self._index: dict[str, set[str]] = defaultdict(set)

    def add_dependency(self, dependent: str, precedent: str) -> None:
        """
        Add a dependency relationship.

        Args:
            dependent: Cell that contains the formula (the cell that depends on precedent)
            precedent: Cell that is referenced in the formula

        Example:
            >>> index = ReverseIndex()
            >>> index.add_dependency("B1", "A1")  # B1 depends on A1
            >>> index.get_dependents("A1")
            ['B1']
        """
        self._index[precedent].add(dependent)

    def add_dependencies(self, dependent: str, precedents: list[str]) -> None:
        """
        Add multiple dependency relationships for a single dependent.

        Args:
            dependent: Cell that contains the formula
            precedents: List of cells referenced in the formula

        Example:
            >>> index = ReverseIndex()
            >>> index.add_dependencies("C1", ["A1", "B1"])  # C1 = A1 + B1
            >>> sorted(index.get_precedents())
            ['A1', 'B1']
        """
        for precedent in precedents:
            self.add_dependency(dependent, precedent)

    def get_dependents(self, precedent: str) -> list[str]:
        """
        Get all cells that depend on a given precedent.

        Args:
            precedent: Cell to look up

        Returns:
            Sorted list of dependent cell addresses (deterministic order)

        Example:
            >>> index = ReverseIndex()
            >>> index.add_dependency("B1", "A1")
            >>> index.add_dependency("C1", "A1")
            >>> index.get_dependents("A1")
            ['B1', 'C1']
        """
        return sorted(self._index.get(precedent, set()))

    def get_precedents(self) -> list[str]:
        """
        Get all precedent cells in the index.

        Returns:
            Sorted list of all precedent addresses

        Example:
            >>> index = ReverseIndex()
            >>> index.add_dependency("B1", "A1")
            >>> index.add_dependency("C1", "A2")
            >>> index.get_precedents()
            ['A1', 'A2']
        """
        return sorted(self._index.keys())

    def has_precedent(self, precedent: str) -> bool:
        """
        Check if a precedent exists in the index.

        Args:
            precedent: Cell address to check

        Returns:
            True if the precedent has at least one dependent
        """
        return precedent in self._index

    def get_all_relationships(self) -> dict[str, list[str]]:
        """
        Get all precedent → dependents relationships.

        Returns:
            Dictionary mapping precedent addresses to sorted lists of dependents

        Example:
            >>> index = ReverseIndex()
            >>> index.add_dependency("B1", "A1")
            >>> index.add_dependency("C1", "A1")
            >>> index.get_all_relationships()
            {'A1': ['B1', 'C1']}
        """
        return {precedent: sorted(dependents) for precedent, dependents in self._index.items()}

    def merge(self, other: "ReverseIndex") -> None:
        """
        Merge another reverse index into this one.

        Args:
            other: Another ReverseIndex to merge

        Example:
            >>> index1 = ReverseIndex()
            >>> index1.add_dependency("B1", "A1")
            >>> index2 = ReverseIndex()
            >>> index2.add_dependency("C1", "A1")
            >>> index1.merge(index2)
            >>> index1.get_dependents("A1")
            ['B1', 'C1']
        """
        for precedent, dependents in other._index.items():
            self._index[precedent].update(dependents)

    def to_dict(self) -> dict[str, list[str]]:
        """
        Export index to dictionary format (for JSON serialization).

        Returns:
            Dictionary with sorted keys and sorted dependent lists
        """
        return {
            precedent: sorted(dependents) for precedent, dependents in sorted(self._index.items())
        }

    @classmethod
    def from_dict(cls, data: dict[str, list[str]]) -> "ReverseIndex":
        """
        Create ReverseIndex from dictionary.

        Args:
            data: Dictionary mapping precedents to lists of dependents

        Returns:
            New ReverseIndex instance
        """
        index = cls()
        for precedent, dependents in data.items():
            for dependent in dependents:
                index.add_dependency(dependent, precedent)
        return index

    def __len__(self) -> int:
        """Return number of precedents in the index"""
        return len(self._index)

    def __repr__(self) -> str:
        """String representation for debugging"""
        return f"ReverseIndex({len(self)} precedents)"
