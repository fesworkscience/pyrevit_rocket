# -*- coding: utf-8 -*-
from pyrevit import forms


def show_short_error(message):
    forms.alert(message, title="GIP Tasks")


def log_exception(message):
    try:
        from . import logger
        logger.exception(message)
    except Exception:
        pass


def run_create_task(uidoc):
    try:
        from . import config, revit_context, task_models, task_service
        from . import marker_service
        from .ui.create_task_window import CreateTaskWindow

        doc = uidoc.Document
        settings = config.load()
        ctx = revit_context.collect(uidoc)
        win = CreateTaskWindow(settings, ctx, uidoc)
        if not win.show_dialog():
            return None
        marker = win.selected_marker
        if not marker:
            show_short_error(u"Сначала выберите маркер задания семейства CPSK_Маркер задания на активном виде.")
            return None
        payload = task_models.new_task_payload(config.load(), win.result, win.revit_context)
        payload, status = task_service.create_task(payload)
        marker_service.write_task_to_marker(doc, marker, payload)
        forms.alert(status, title="GIP Tasks", warn_icon=False)
        return payload
    except Exception:
        log_exception("Create task command failed")
        show_short_error(u"Не удалось передать задание. Подробности записаны в лог.")
        return None


def run_journal(uidoc):
    try:
        from .ui.task_journal_window import TaskJournalWindow

        win = TaskJournalWindow(uidoc)
        win.show_dialog()
    except Exception:
        log_exception("Journal command failed")
        show_short_error(u"Не удалось открыть журнал. Подробности записаны в лог.")
