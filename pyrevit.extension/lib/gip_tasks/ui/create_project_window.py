# -*- coding: utf-8 -*-
import os
from pyrevit import forms
from gip_tasks import config
from gip_tasks.ui import theme


def xaml_path(name):
    return os.path.join(config.extension_root(), "lib", "gip_tasks", "ui", name)


class CreateProjectWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, xaml_path("create_project_window.xaml"))
        self.result = None
        theme.apply_theme(self)

    def create_click(self, sender, args):
        name = (self.NameBox.Text or "").strip()
        if not name:
            forms.alert(u"Укажите название проекта", title="GIP Tasks")
            return
        self.result = {
            "name": name,
            "is_active": True,
        }
        self.DialogResult = True
        self.Close()

    def cancel_click(self, sender, args):
        self.DialogResult = False
        self.Close()
