"""
Configuration module for the PO auto-scan tool.

This module now exposes a ``Settings`` class (see improvement.txt item 7) to
centralize configuration values. An instance of ``Settings`` is created at
import time and module-level constants are kept for backward compatibility.
If you wish to override settings at runtime, instantiate your own ``Settings``
object or modify the attributes on ``settings``.
"""

from pathlib import Path


class Settings:
    """Container for configurable constants used across the project."""

    # Base working directory for output files
    BASE_DIR: Path = Path.home() / "Desktop" / "Auto Tools" / "Scanned PO"
    TEMP_DIR: Path = BASE_DIR / "temp"
    FILTERED_DIR: Path = BASE_DIR / "PO_Filtered"
    LOG_DIR: Path = BASE_DIR / "log"

    # Template file names for attachments
    TEMPLATE_LOCAL: Path = TEMP_DIR / "1_LOCAL HS code request.xlsx"
    TEMPLATE_OVERSEA: Path = TEMP_DIR / "2_OVERSEA Machine list.xlsx"
    NON_CDS_SUPPLIER_FILE: Path = TEMP_DIR / "Non-CDs Supplier.csv"

    # Concurrency settings
    MAX_WORKERS: int = 4

    def __init__(self, **overrides):
        """
        Optionally override configuration values via keyword arguments.

        Example:
        >>> cfg = Settings(BASE_DIR=Path("/tmp/po_scan"), MAX_WORKERS=8)
        """
        for key, value in overrides.items():
            if hasattr(self, key):
                setattr(self, key, value)


# Default settings instance
settings = Settings()

# Backwards compatibility: expose selected attributes at module level
BASE_DIR = settings.BASE_DIR
TEMP_DIR = settings.TEMP_DIR
FILTERED_DIR = settings.FILTERED_DIR
LOG_DIR = settings.LOG_DIR
TEMPLATE_LOCAL = settings.TEMPLATE_LOCAL
TEMPLATE_OVERSEA = settings.TEMPLATE_OVERSEA
NON_CDS_SUPPLIER_FILE = settings.NON_CDS_SUPPLIER_FILE
MAX_WORKERS = settings.MAX_WORKERS
