"""Where the app's own files are: the app folder (templates, static, data, projects, logs), wherever the code sits.

Modules live in feature folders under engine/ (engine/core/, engine/courier/, engine/addresses/ ...), so a path worked out from a module's own file would
point inside that folder. Everything asks here instead. A built .exe keeps its files next to the .exe.
"""
import os
import sys

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DATA_DIR = os.path.join(BASE_DIR, 'data')
