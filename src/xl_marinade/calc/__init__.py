"""Experimental workbook recalculator (``marinade calc``).

Importing this package never imports ``ironcalc``; only
:mod:`xl_marinade.calc.engine` does, lazily, when an engine is loaded.
"""

from xl_marinade.calc.values import ValuesRequest, values

__all__ = ["ValuesRequest", "values"]
