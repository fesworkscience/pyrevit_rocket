# -*- coding: utf-8 -*-
import os
try:
    unicode
except NameError:
    unicode = str

try:
    from System import DateTime
except Exception:
    DateTime = None

from System.Windows import Visibility
from pyrevit import forms
from gip_tasks import config, marker_service, project_service, revit_context, task_models
from gip_tasks.ui import theme
from gip_tasks.ui.create_project_window import CreateProjectWindow
from gip_tasks.ui.select_project_window import ProjectItem


def xaml_path(name):
    return os.path.join(config.extension_root(), "lib", "gip_tasks", "ui", name)


def _visible(value):
    return Visibility.Visible if value else Visibility.Collapsed


class CreateTaskWindow(forms.WPFWindow):
    def __init__(self, settings, revit_context_data, uidoc=None):
        forms.WPFWindow.__init__(self, xaml_path("create_task_window.xaml"))
        self.settings = settings
        self.uidoc = uidoc
        self.revit_context = revit_context_data or {}
        self.selected_marker = None
        self.result = None
        self.projects = []
        self._bind()
        theme.apply_theme(self)
        self._update_marker_card()
        self._set_task_fields_enabled(bool((self.revit_context.get("marker") or {}).get("element_id")))
        self.update_send_state()

    def _bind(self):
        self.CreateProjectButton.Visibility = _visible(config.can_create_project(self.settings))
        self.SenderDisciplineBox.ItemsSource = task_models.DISCIPLINES
        self.ReceiverDisciplineBox.ItemsSource = task_models.DISCIPLINES
        self.TypeBox.ItemsSource = task_models.TASK_TYPES
        self.PriorityBox.ItemsSource = task_models.PRIORITIES
        self.SenderDisciplineBox.SelectedItem = self.settings.get("CURRENT_DISCIPLINE") or u"КЖ"
        self.ReceiverDisciplineBox.SelectedIndex = 1 if len(task_models.DISCIPLINES) > 1 else 0
        self.TypeBox.SelectedItem = u"прочее"
        self.PriorityBox.SelectedItem = u"обычный"
        self.refresh_projects(silent=True, use_server=False)
        if (self.revit_context.get("marker") or {}).get("element_id"):
            self._set_marker_context(None, self.revit_context)

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

    def _select_project_in_combo(self):
        current_id = self.settings.get("CURRENT_PROJECT_ID")
        for item in self.ProjectCombo.ItemsSource or []:
            if item.data.get("id") == current_id:
                self.ProjectCombo.SelectedItem = item
                break

    def refresh_projects(self, silent=False, use_server=True):
        projects, status = project_service.load_projects(use_server=use_server)
        self.projects = projects
        self.ProjectCombo.ItemsSource = [ProjectItem(x) for x in projects]
        self._select_project_in_combo()
        mode_text = self._mode_text(status)
        self.ConnectionStatusText.Text = mode_text
        self.ModeBadge.Text = mode_text
        self.StatusText.Text = status or u"Готово"
        theme.set_connection_state(self.ConnectionDot, mode_text)
        self.update_send_state()

    def project_changed(self, sender, args):
        item = self.ProjectCombo.SelectedItem
        if item:
            self.settings = project_service.choose_project(item.data)
        self.update_send_state()

    def refresh_projects_click(self, sender, args):
        self.refresh_projects(use_server=True)

    def create_project_click(self, sender, args):
        win = CreateProjectWindow()
        if win.show_dialog():
            payload = task_models.new_project_payload(config.load(), win.result)
            project, status = project_service.create_project(payload)
            self.settings = project_service.choose_project(project)
            self.refresh_projects(silent=True, use_server=False)
            self.StatusText.Text = status

    def _set_marker_context(self, marker, context):
        self.selected_marker = marker
        self.revit_context = context or self.revit_context or {}
        self._update_marker_card()
        self._set_task_fields_enabled(bool((self.revit_context.get("marker") or {}).get("element_id")))
        self.update_send_state()

    def pick_marker_click(self, sender, args):
        if not self.uidoc:
            forms.alert(u"Активный документ Revit недоступен", title="GIP Tasks")
            return
        marker = None
        message = u""
        try:
            try:
                self.Hide()
            except Exception:
                pass
            marker, message = marker_service.pick_marker(self.uidoc)
        finally:
            try:
                self.Show()
                self.Activate()
            except Exception:
                pass
        if message:
            forms.alert(u"Выберите маркер задания семейства CPSK_Маркер задания на активном виде.", title="GIP Tasks")
            self.StatusText.Text = message
            return
        if not marker:
            return
        context = revit_context.collect(self.uidoc, marker)
        params = marker_service.marker_parameter_values(marker)
        context.setdefault("marker", {})["parameters"] = params
        self._set_marker_context(marker, context)
        self.StatusText.Text = u"Маркер выбран"

    def _marker_value(self, key, fallback=u"не задано"):
        marker = self.revit_context.get("marker") or {}
        value = marker.get(key)
        return unicode(value) if value not in (None, "") else fallback

    def _update_marker_card(self):
        marker = self.revit_context.get("marker") or {}
        active_view = self.revit_context.get("active_view") or {}
        has_marker = bool(marker.get("element_id"))
        self.MarkerSelectedPanel.Visibility = _visible(has_marker)
        self.MarkerDetailsPanel.Visibility = _visible(has_marker)
        self.MarkerHintText.Visibility = _visible(not has_marker)
        self.MarkerNameText.Text = marker.get("family_name") or u"CPSK_Маркер задания"
        self.MarkerTypeText.Text = marker.get("type_name") or u"не задано"
        self.MarkerViewText.Text = active_view.get("name") or u"не задано"
        self.MarkerLevelText.Text = active_view.get("level") or u"не задано"
        self.MarkerIdText.Text = unicode(marker.get("element_id") or u"не задано")
        params = marker.get("parameters") or {}
        lines = []
        for key in sorted(params.keys()):
            value = params.get(key) or u"не задано"
            lines.append(u"%s: %s" % (key, value))
        self.MarkerParamsText.Text = u"\n".join(lines) if lines else u"Параметры маркера не заполнены"

    def _task_field_controls(self):
        return [
            self.SenderDisciplineBox,
            self.ReceiverDisciplineBox,
            self.TypeBox,
            self.PriorityBox,
            self.DueDatePicker,
            self.TodayButton,
            self.TomorrowButton,
            self.ThreeDaysButton,
            self.SevenDaysButton,
            self.DescriptionBox,
            self.CommentBox,
        ]

    def _set_task_fields_enabled(self, enabled):
        for control in self._task_field_controls():
            try:
                control.IsEnabled = enabled
            except Exception:
                pass

    def _due_date_text(self):
        selected = self.DueDatePicker.SelectedDate
        if selected is None:
            return u""
        try:
            selected = selected.Value
        except Exception:
            pass
        try:
            return selected.ToString("dd.MM.yyyy")
        except Exception:
            return unicode(selected)

    def set_due_date(self, days):
        if DateTime is None:
            return
        self.DueDatePicker.SelectedDate = DateTime.Today.AddDays(days)
        self.update_send_state()

    def today_click(self, sender, args):
        self.set_due_date(0)

    def tomorrow_click(self, sender, args):
        self.set_due_date(1)

    def three_days_click(self, sender, args):
        self.set_due_date(3)

    def seven_days_click(self, sender, args):
        self.set_due_date(7)

    def field_changed(self, sender=None, args=None):
        self._update_placeholders()
        self.update_send_state()

    def _update_placeholders(self):
        try:
            self.DescriptionPlaceholder.Visibility = _visible(not bool((self.DescriptionBox.Text or "").strip()))
            self.CommentPlaceholder.Visibility = _visible(not bool((self.CommentBox.Text or "").strip()))
        except Exception:
            pass

    def _update_due_chips(self):
        selected = self._due_date_text()
        buttons = [
            (self.TodayButton, 0),
            (self.TomorrowButton, 1),
            (self.ThreeDaysButton, 3),
            (self.SevenDaysButton, 7),
        ]
        for button, days in buttons:
            active = False
            if DateTime is not None and selected:
                try:
                    active = DateTime.Today.AddDays(days).ToString("dd.MM.yyyy") == selected
                except Exception:
                    active = False
            theme.set_chip_state(button, active)

    def _validation_errors(self):
        errors = []
        if not self.settings.get("CURRENT_PROJECT_ID"):
            errors.append(u"выберите проект")
        if not bool((self.revit_context.get("marker") or {}).get("element_id")):
            errors.append(u"выберите маркер задания")
        if not self.SenderDisciplineBox.SelectedItem:
            errors.append(u"укажите раздел «От кого»")
        if not self.ReceiverDisciplineBox.SelectedItem:
            errors.append(u"укажите раздел «Кому»")
        if not self.TypeBox.SelectedItem:
            errors.append(u"выберите тип")
        if not self.PriorityBox.SelectedItem:
            errors.append(u"выберите приоритет")
        if not self._due_date_text():
            errors.append(u"выберите срок")
        if not (self.DescriptionBox.Text or "").strip():
            errors.append(u"заполните описание")
        return errors

    def update_send_state(self):
        try:
            errors = self._validation_errors()
            self.SendButton.IsEnabled = len(errors) == 0
            self.ValidationText.Text = u"Готово к передаче" if not errors else u"Для передачи заполните: " + u", ".join(errors)
            self._update_placeholders()
            self._update_due_chips()
        except Exception:
            pass

    def send_click(self, sender, args):
        item = self.ProjectCombo.SelectedItem
        if item:
            self.settings = project_service.choose_project(item.data)
        errors = self._validation_errors()
        if errors:
            self.ValidationText.Text = u"Заполните обязательные поля: " + u", ".join(errors)
            forms.alert(self.ValidationText.Text, title="GIP Tasks")
            return
        self.result = {
            "sender_discipline": self.SenderDisciplineBox.SelectedItem,
            "receiver_discipline": self.ReceiverDisciplineBox.SelectedItem,
            "type": self.TypeBox.SelectedItem,
            "priority": self.PriorityBox.SelectedItem,
            "due_date": self._due_date_text(),
            "description": (self.DescriptionBox.Text or "").strip(),
            "comment": (self.CommentBox.Text or "").strip(),
            "marker_status": u"Новое",
        }
        self.DialogResult = True
        self.Close()

    def cancel_click(self, sender, args):
        self.DialogResult = False
        self.Close()
