"""pdf2vector: recebe um PDF, extrai e vetoriza todo o conteúdo e gera um Excel."""

from .extractor import DocumentData, extract
from .excel import build_workbook, write_excel
from .pipeline import converter
from .vectorizer import Vetorizacao, vectorize

__all__ = ["DocumentData", "Vetorizacao", "build_workbook", "converter", "extract",
           "vectorize", "write_excel"]
__version__ = "0.1.0"
