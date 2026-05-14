# -*- coding: utf-8 -*-
import os
import traceback
from datetime import datetime

from . import config


def log_path():
    return os.path.join(config.appdata_dir(), "gip_tasks.log")


def write(message):
    try:
        line = "[%s] %s\n" % (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), message)
        with open(log_path(), "ab") as fp:
            fp.write(line.encode("utf-8"))
    except Exception:
        pass


def exception(message):
    write("%s\n%s" % (message, traceback.format_exc()))

