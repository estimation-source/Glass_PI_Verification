# ==========================================================
# GLASS PI VERIFICATION SYSTEM
# CUSTOM EXCEPTIONS MODULE
# ==========================================================

class GlassPIException(Exception):
    """Base exception class for Glass PI project."""
    pass

class ExcelReaderError(GlassPIException):
    """Raised when Excel file reading or parsing fails."""
    pass

class PDFReaderError(GlassPIException):
    """Raised when PDF file reading or parsing fails."""
    pass

class MatchingError(GlassPIException):
    """Raised when verification/matching logic fails."""
    pass

class ReportError(GlassPIException):
    """Raised when generating output report fails."""
    pass

class ReportGenerationError(GlassPIException):
    """Raised when report generation fails in report.py."""
    pass

class ConfigError(GlassPIException):
    """Raised when config loading fails."""
    pass


class PDFReadError(Exception):
    """Custom exception for PDF Reading failures"""
    pass

    