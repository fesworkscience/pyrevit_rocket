# -*- coding: utf-8 -*-
"""Открывает мастер передачи задания между разделами."""

__title__ = "Передать\nзадание"
__author__ = "CPSK"

import os
import sys


SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from gip_tasks import commands


commands.run_create_task(__revit__.ActiveUIDocument)
