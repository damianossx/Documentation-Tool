"""
SmartDocs Insight package.

Provides a Tkinter-based desktop assistant for:
- Analyzing invoice CSV/PDF combinations,
- Classifying items by COO (EU vs non-EU),
- Calculating non-EU weight,
- Generating Excel tracking sheets for Certificate of Origin requests.

The primary entrypoint for the GUI is `show_gui()` exposed from `main.py`.
"""

from .main import show_gui

__all__ = ["show_gui"]

# Semantic version of the application
__version__ = "2.2.0"