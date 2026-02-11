"""
プロジェクト管理MPCサーバー

Claude DesktopからExcelファイルのプロジェクト管理を行うMPCサーバーです。
"""

__version__ = "1.0.0"
__author__ = "Your Name"

from .server import main
from .excel_handler import ExcelHandler
from .config import ConfigManager

__all__ = [
    "main",
    "ExcelHandler",
    "ConfigManager",
]
