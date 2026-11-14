"""PPTX Citation Manager."""
__version__ = "1.0.0"

from .pptx_cite import BibTeXParser, CitationProcessor, load_config

__all__ = ["BibTeXParser", "CitationProcessor", "load_config"]
