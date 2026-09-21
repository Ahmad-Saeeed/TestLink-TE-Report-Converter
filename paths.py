"""Resource path resolution that works both as a plain script and as a
frozen PyInstaller app (onedir or onefile — sys._MEIPASS is set in both
cases, pointing at the app's own folder for onedir or a temp extraction
folder for onefile)."""
import sys
import os


def resource_path(relative_path):
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)
