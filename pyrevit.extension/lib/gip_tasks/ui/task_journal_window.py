# -*- coding: utf-8 -*-
import os
try:
    unicode
except NameError:
    unicode = str
from datetime import datetime

from System.Windows import Visibility
from System.Collections.Generic import List
from Autodesk.Revit.DB import ElementId
from pyrevit import forms
from gip_tasks import config, local_cache, marker_service, project_service, task_models, task_service
from gip_tasks.ui import theme
from gip_tasks.ui.create_project_window import CreateProjectWindow
from gip_tasks.ui.select_project_window import ProjectItem


def xaml_path(name):
    return os.path.join(config.extension_root(), "lib", "gip_tasks", "ui", name)


def _visible(value):
    return Visibility.Visible if value else Visibility.Collapsed


def _lower(value):
    return (value or u"").lower()


def _parse_due_date(value):
    value = (value or u"").strip()
    if not value:
        return None
    for fmt in ("%d.%m.%Y", "%d.%m.%Y %H:%M", "%Y-%m-%d", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(value[:len(datetime.now().strftime(fmt))], fmt).date()
        except Exception:
            pass
    return None


def _is_overdue(task):
    due_date = _parse_due_date(task.get("due_date"))
    if not due_date:
        return False
    marker_status = task.get("marker_status") or task_models.server_to_marker_status(task.get("status"))
    if marker_status in (u"Отменено", u"Выполнено"):
        return False
    if task.get("status") in ("closed", "done", "completed", "cancelled"):
        return False
    return due_date < datetime.now().date()


def display_status(task):
    marker_status = task.get("marker_status") or task_models.server_to_marker_status(task.get("status"))
    server_status = task.get("status")
    if marker_status == u"Отменено" or server_status in ("cancelled", "returned"):
        return u"Отменено"
    if server_status in ("closed", "done", "completed", u"Выполнено"):
        return u"Выполнено"
    if _is_overdue(task):
        return u"Просрочено"
    if marker_status == u"Принято" or server_status in ("in_progress", "review"):
        return u"В работе"
    return u"Ожидает"


def priority_label(value):
    value = value or u""
    lowered = value.lower()
    if lowered == u"срочный":
        return u"Срочный"
    if lowered == u"низкий":
        return u"Низкий"
    if lowered == u"высокий":
        return u"Высокий"
    return u"Обычный"


def priority_colors(value):
    label = priority_label(value)
    if label == u"Срочный":
        return "#FEE4E2", "#B42318"
    if label == u"Низкий":
        return "#ECFDF3", "#027A48"
    if label == u"Высокий":
        return "#FEF3C7", "#92400E"
    return "#E0EAFF", "#3136FF"


def status_colors(value):
    if value == u"Просрочено":
        return "#FEE4E2", "#B42318"
    if value == u"Отменено":
        return "#F2F4F7", "#475467"
    if value == u"Выполнено":
        return "#ECFDF3", "#027A48"
    if value == u"В работе":
        return "#E0EAFF", "#3136FF"
    return "#FEF0C7", "#B54708"


def project_display_text(task):
    code = task.get("project_code") or ""
    name = task.get("project_name") or task.get("project_id") or ""
    return ("%s - %s" % (code, name)).strip(" -")


class TaskRow(object):
    def __init__(self, data):
        self.data = data
        tid = data.get("id") or data.get("temporary_id") or ""
        self.Number = tid[:8]
        self.Route = u"%s → %s" % (data.get("sender_discipline") or "", data.get("receiver_discipline") or "")
        self.Type = data.get("type") or u"Задание"
        self.DueDate = data.get("due_date") or u""
        self.Priority = priority_label(data.get("priority"))
        self.Status = display_status(data)
        self.PriorityBackground, self.PriorityForeground = priority_colors(data.get("priority"))
        self.StatusBackground, self.StatusForeground = status_colors(self.Status)
        self.RowStripBrush = "#E2E5F0"
        self.Level = data.get("level_name") or ""
        self.View = data.get("view_name") or ""
        self.Description = (data.get("description") or "").replace("\n", " ")[:140]


class TaskJournalWindow(forms.WPFWindow):
    def __init__(self, uidoc):
        forms.WPFWindow.__init__(self, xaml_path("task_journal_window.xaml"))
        self.uidoc = uidoc
        self.settings = config.load()
        self.tasks = []
        self.selected_task = None
        self.active_quick_tab = "all"
        self._bind_filters()
        theme.apply_theme(self)
        self.refresh_projects(use_server=False)
        self.load_tasks(use_server=False)
        self._set_active_tab_visuals()
        self._show_card()
        try:
            def _loaded(sender, args):
                self._set_active_tab_visuals()
                self._update_search_placeholder()
                self._show_card()
            self.Loaded += _loaded
            self._gip_journal_loaded_handler = _loaded
        except Exception:
            pass

    def _bind_filters(self):
        disciplines = [u"Все разделы"] + task_models.DISCIPLINES
        self.SenderFilter.ItemsSource = disciplines
        self.ReceiverFilter.ItemsSource = disciplines
        self.StatusFilter.ItemsSource = [u"Все статусы", u"Ожидает", u"В работе", u"Выполнено", u"Отменено", u"Просрочено"]
        self.PriorityFilter.ItemsSource = [u"Все"] + task_models.PRIORITIES
        self.SenderFilter.SelectedIndex = 0
        self.ReceiverFilter.SelectedIndex = 0
        self.StatusFilter.SelectedIndex = 0
        self.PriorityFilter.SelectedIndex = 0
        self.SearchBox.Text = ""
        self.CreateProjectButton.Visibility = _visible(config.can_create_project(self.settings))

    def _mode_text(self, status):
        status = status or u""
        lowered = status.lower()
        if config.is_local_mode(self.settings):
            return u"Локальный режим"
        if u"недоступен" in lowered:
            return u"Сервер недоступен"
        if u"кэш" in lowered:
            return u"Локальный кэш"
        if u"локальный режим" in lowered:
            return u"Локальный режим"
        return u"Сервер подключен"

    def refresh_projects(self, use_server=True):
        projects, status = project_service.load_projects(use_server=use_server)
        self.ProjectCombo.ItemsSource = [ProjectItem(x) for x in projects]
        current_id = self.settings.get("CURRENT_PROJECT_ID")
        for item in self.ProjectCombo.ItemsSource or []:
            if item.data.get("id") == current_id:
                self.ProjectCombo.SelectedItem = item
                break
        mode_text = self._mode_text(status)
        self.ConnectionStatusText.Text = mode_text
        theme.set_connection_state(self.ConnectionDot, mode_text)
        self.StatusText.Text = status or self._queue_status()

    def _queue_status(self):
        project_id = self.settings.get("CURRENT_PROJECT_ID")
        if not project_id:
            return u"Сначала выберите проект"
        count = local_cache.pending_count(project_id)
        return u"Ожидает отправки: %s" % count if count else u""

    def project_changed(self, sender, args):
        item = self.ProjectCombo.SelectedItem
        if item:
            self.settings = project_service.choose_project(item.data)
            self.load_tasks(use_server=False)

    def _matches_quick_tab(self, task):
        discipline = self.settings.get("CURRENT_DISCIPLINE") or u""
        tab = self.active_quick_tab
        if tab == "mine":
            return task.get("receiver_discipline") == discipline
        if tab == "from_me":
            return task.get("sender_discipline") == discipline
        if tab == "urgent":
            return _lower(task.get("priority")) == u"срочный"
        if tab == "overdue":
            return display_status(task) == u"Просрочено"
        if tab == "cancelled":
            return display_status(task) == u"Отменено"
        return True

    def _filtered_rows(self):
        text = (self.SearchBox.Text or "").lower()
        status = self.StatusFilter.SelectedItem
        sender = self.SenderFilter.SelectedItem
        receiver = self.ReceiverFilter.SelectedItem
        priority = self.PriorityFilter.SelectedItem
        rows = []
        for task in self.tasks:
            if not self._matches_quick_tab(task):
                continue
            marker_status = display_status(task)
            if status and status != u"Все статусы" and marker_status != status:
                continue
            if sender and sender != u"Все разделы" and task.get("sender_discipline") != sender:
                continue
            if receiver and receiver != u"Все разделы" and task.get("receiver_discipline") != receiver:
                continue
            if priority and priority != u"Все" and task.get("priority") != priority:
                continue
            hay = u"%s %s %s %s %s" % (
                task.get("description") or "",
                task.get("type") or "",
                task.get("view_name") or "",
                task.get("level_name") or "",
                project_display_text(task),
            )
            if text and text not in hay.lower():
                continue
            rows.append(TaskRow(task))
        return rows

    def _set_empty_state(self, rows, has_project=True):
        self.EmptyStatePanel.Visibility = _visible(has_project and not rows)

    def load_tasks(self, use_server=True):
        project_id = self.settings.get("CURRENT_PROJECT_ID")
        if not project_id:
            self.tasks = []
            self.TasksList.ItemsSource = []
            self.selected_task = None
            self._set_empty_state([], has_project=False)
            self.StatusText.Text = u"Сначала выберите проект"
            self._show_card()
            return
        self.tasks, status = task_service.list_tasks(project_id, use_server=use_server)
        self._apply_filters(status)

    def _apply_filters(self, status_text=None):
        rows = self._filtered_rows()
        self.TasksList.ItemsSource = rows
        self._set_empty_state(rows, has_project=bool(self.settings.get("CURRENT_PROJECT_ID")))
        if status_text is not None:
            self.StatusText.Text = status_text or self._queue_status() or u"Заданий: %s" % len(self.tasks)
        else:
            self.StatusText.Text = self._queue_status() or u"Заданий: %s" % len(rows)
        if self.selected_task:
            keep_id = self.selected_task.get("id") or self.selected_task.get("temporary_id")
            for row in rows:
                row_id = row.data.get("id") or row.data.get("temporary_id")
                if row_id == keep_id:
                    self.TasksList.SelectedItem = row
                    self.selected_task = row.data
                    self._show_card()
                    return
        if rows:
            self.TasksList.SelectedItem = rows[0]
            self.selected_task = rows[0].data
            self._show_card()
            return
        self.selected_task = None
        self._show_card()

    def filter_changed(self, sender, args):
        self._update_search_placeholder()
        self._apply_filters()

    def _update_search_placeholder(self):
        try:
            self.SearchPlaceholder.Visibility = _visible(not bool((self.SearchBox.Text or "").strip()))
        except Exception:
            pass

    def reset_filters_click(self, sender, args):
        self.SearchBox.Text = ""
        self.SenderFilter.SelectedIndex = 0
        self.ReceiverFilter.SelectedIndex = 0
        self.StatusFilter.SelectedIndex = 0
        self.PriorityFilter.SelectedIndex = 0
        self.active_quick_tab = "all"
        self._set_active_tab_visuals()
        self._update_search_placeholder()
        self._apply_filters()

    def quick_tab_click(self, sender, args):
        try:
            self.active_quick_tab = str(sender.CommandParameter)
        except Exception:
            self.active_quick_tab = "all"
        self._set_active_tab_visuals()
        self._apply_filters()

    def _tab_buttons(self):
        return {
            "all": self.TabAllButton,
            "mine": self.TabMineButton,
            "from_me": self.TabFromMeButton,
            "urgent": self.TabUrgentButton,
            "overdue": self.TabOverdueButton,
            "cancelled": self.TabCancelledButton,
        }

    def _set_active_tab_visuals(self):
        for key, button in self._tab_buttons().items():
            theme.set_quick_tab_state(button, key == self.active_quick_tab)

    def create_project_click(self, sender, args):
        win = CreateProjectWindow()
        if win.show_dialog():
            payload = task_models.new_project_payload(config.load(), win.result)
            project, status = project_service.create_project(payload)
            self.settings = project_service.choose_project(project)
            self.refresh_projects(use_server=False)
            self.load_tasks(use_server=False)
            self.StatusText.Text = status

    def new_task_click(self, sender, args):
        from gip_tasks import commands
        payload = commands.run_create_task(self.uidoc)
        if payload:
            self.settings = config.load()
            self.refresh_projects(use_server=False)
            self.load_tasks(use_server=False)

    def refresh_click(self, sender, args):
        if not self.settings.get("CURRENT_PROJECT_ID"):
            forms.alert(u"Сначала выберите проект", title="GIP Tasks")
            return
        sent = project_service.sync_pending(self.settings.get("CURRENT_PROJECT_ID"))
        self.refresh_projects(use_server=True)
        self.load_tasks(use_server=True)
        self.StatusText.Text = u"Синхронизировано" if sent else (self.StatusText.Text or u"Синхронизировано")

    def task_selected(self, sender, args):
        row = self.TasksList.SelectedItem
        self.selected_task = row.data if row else None
        self._show_card()

    def _set_card_badges(self, status, priority):
        bg, fg = status_colors(status)
        theme.set_badge(self.CardStatusBadge, self.CardStatusText, bg, fg)
        bg, fg = priority_colors(priority)
        theme.set_badge(self.CardPriorityBadge, self.CardPriorityText, bg, fg)

    def _show_card(self):
        task = self.selected_task
        if not task:
            self.CardEmptyStatePanel.Visibility = Visibility.Visible
            self.CardDetailsPanel.Visibility = Visibility.Collapsed
            self.CardActionsPanel.Visibility = Visibility.Collapsed
            self.CardType.Text = u"Выберите задание"
            self.CardStatusText.Text = u""
            self.CardProject.Text = u""
            self.CardRoute.Text = u""
            self.CardDueDate.Text = u""
            self.CardPriorityText.Text = u""
            self.CardView.Text = u""
            self.CardLevel.Text = u""
            self.CardDescription.Text = u""
            self.CardMarker.Text = u""
            self.CommentBox.Text = u""
            self.CardMarkerParams.Text = u""
            self.AcceptButton.IsEnabled = False
            self.CancelTaskButton.IsEnabled = False
            self.FindMarkerButton.IsEnabled = False
            return
        self.CardEmptyStatePanel.Visibility = Visibility.Collapsed
        self.CardDetailsPanel.Visibility = Visibility.Visible
        self.CardActionsPanel.Visibility = Visibility.Visible
        status = display_status(task)
        marker_name = task.get("marker_family_name") or u"CPSK_Маркер задания"
        marker_id = task.get("marker_element_id") or u""
        route = u"%s → %s" % (task.get("sender_discipline") or "", task.get("receiver_discipline") or "")
        self.CardType.Text = task.get("type") or u"Задание"
        self.CardStatusText.Text = status
        self.CardProject.Text = u"Проект: %s" % (project_display_text(task) or self.settings.get("CURRENT_PROJECT_NAME") or u"не задан")
        self.CardRoute.Text = route
        self.CardDueDate.Text = u"Срок: %s" % (task.get("due_date") or u"не задан")
        self.CardPriorityText.Text = priority_label(task.get("priority"))
        self.CardView.Text = task.get("view_name") or u"не задано"
        self.CardLevel.Text = task.get("level_name") or u"не задано"
        self.CardDescription.Text = task.get("description") or u""
        self.CommentBox.Text = task.get("comment") or u""
        self._update_card_comment_placeholder()
        self.CardMarker.Text = u"%s · ID %s" % (marker_name, marker_id) if marker_id else marker_name
        self.CardMarkerParams.Text = u"Параметры маркера обновляются при записи задания в семейство."
        self._set_card_badges(status, task.get("priority"))
        inactive = status in (u"Отменено", u"Выполнено")
        self.AcceptButton.IsEnabled = not inactive
        self.CancelTaskButton.IsEnabled = not inactive
        self.FindMarkerButton.IsEnabled = bool(task.get("marker_unique_id"))

    def _update_card_comment_placeholder(self):
        try:
            self.CardCommentPlaceholder.Visibility = _visible(not bool((self.CommentBox.Text or "").strip()))
        except Exception:
            pass

    def comment_text_changed(self, sender, args):
        self._update_card_comment_placeholder()

    def _selected_or_warn(self):
        if not self.selected_task:
            forms.alert(u"Выберите задание", title="GIP Tasks")
            return None
        return self.selected_task

    def _update_marker_for_task(self, task):
        marker = marker_service.find_marker_by_unique_id(self.uidoc.Document, task.get("marker_unique_id"))
        if marker:
            marker_service.write_task_to_marker(self.uidoc.Document, marker, task, update_date=False)

    def accept_click(self, sender, args):
        task = self._selected_or_warn()
        if not task:
            return
        if display_status(task) in (u"Отменено", u"Выполнено"):
            forms.alert(u"Для завершенного или отмененного задания действие недоступно", title="GIP Tasks")
            return
        self.selected_task, status = task_service.update_status(task, u"Принято")
        self._update_marker_for_task(self.selected_task)
        self.load_tasks(use_server=False)
        self.StatusText.Text = status

    def cancel_task_click(self, sender, args):
        task = self._selected_or_warn()
        if not task:
            return
        if display_status(task) in (u"Отменено", u"Выполнено"):
            forms.alert(u"Для завершенного или отмененного задания действие недоступно", title="GIP Tasks")
            return
        self.selected_task, status = task_service.update_status(task, u"Отменено")
        self._update_marker_for_task(self.selected_task)
        self.load_tasks(use_server=False)
        self.StatusText.Text = status

    def comment_click(self, sender, args):
        task = self._selected_or_warn()
        if not task:
            return
        self.selected_task, status = task_service.add_comment(task, self.CommentBox.Text or "")
        self._update_marker_for_task(self.selected_task)
        self.load_tasks(use_server=False)
        self.StatusText.Text = status

    def find_marker_click(self, sender, args):
        task = self._selected_or_warn()
        if not task:
            return
        marker = marker_service.find_marker_by_unique_id(self.uidoc.Document, task.get("marker_unique_id"))
        if not marker:
            forms.alert(u"Маркер не найден в текущей модели", title="GIP Tasks")
            return
        ids = List[ElementId]()
        ids.Add(marker.Id)
        self.uidoc.Selection.SetElementIds(ids)
        forms.alert(u"Маркер выбран в модели", title="GIP Tasks", warn_icon=False)
