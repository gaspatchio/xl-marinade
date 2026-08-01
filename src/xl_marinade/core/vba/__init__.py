# ABOUTME: VBA extraction package — grammar-based parser + procedure extractor.
# ABOUTME: Uses antlr4-vba (vba_ccParser) for tokenization; builds structured VBAExtraction.

from xl_marinade.core.vba.extractor import (
    VBADeclaration,
    VBAExtraction,
    VBAModule,
    VBAProcedure,
    extract_vba,
)

__all__ = ["extract_vba", "VBAExtraction", "VBAModule", "VBAProcedure", "VBADeclaration"]
