# -*- coding: utf-8 -*-
import os
from System.Windows import Visibility
from pyrevit import forms
from gip_tasks import config, project_service, task_models
from gip_tasks.ui import theme
from gip_tasks.ui.create_project_window import CreateProjectWindow


def xaml_path(name):
    return os.path.join(config.extension_root(), "lib", "gip_tasks", "ui", name)


def project_display(project):
    code = project.get("code") or project.get("project_code") or ""
    name = project.get("name") or project.get("project_name") or project.get("id") or ""
    label = ("%s - %s" % (code, name)).strip(" -")
    if project.get("is_local"):
        label += u" (локальный проект)"
    return label


class ProjectItem(object):
    def __init__(self, data):
        self.data = data
        self.display_name = project_display(data)


class SelectProjectWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, xaml_path("select_project_window.xaml"))
        self.selected_project = None
        self.CreateProjectButton.Visibility = Visibility.Visible if config.can_create_project() else Visibility.Collapsed
        theme.apply_theme(self)
        self.refresh(use_server=False)

    def refresh(self, use_server=True):
        projects, status = project_service.load_projects(use_server=use_server)
        self.ProjectsList.ItemsSource = [ProjectItem(x) for x in projects]
        self.StatusText.Text = status or (u"Проектов: %s" % len(projects))

    def refresh_click(self, sender, args):
        self.refresh(use_server=True)

    def create_project_click(self, sender, args):
        if not config.can_create_project():
            forms.alert(u"У текущей роли нет права создавать проекты", title="GIP Tasks")
            return
        win = CreateProjectWindow()
        if win.show_dialog():
            payload = task_models.new_project_payload(config.load(), win.result)
            project, status = project_service.create_project(payload)
            self.selected_project = project
            self.StatusText.Text = status
            self.DialogResult = True
            self.Close()

    def select_click(self, sender, args):
        item = self.ProjectsList.SelectedItem
        if not item:
            forms.alert(u"Выберите проект", title="GIP Tasks")
            return
        self.selected_project = item.data
        self.DialogResult = True
        self.Close()

    def cancel_click(self, sender, args):
        self.DialogResult = False
        self.Close()
