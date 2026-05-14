# -*- coding: utf-8 -*-
try:
    unicode
except NameError:
    unicode = str

from datetime import datetime

try:
    from Autodesk.Revit.DB import BuiltInCategory, ElementId, StorageType, Transaction
    from Autodesk.Revit.Exceptions import OperationCanceledException
    from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
except Exception:
    BuiltInCategory = ElementId = StorageType = Transaction = None
    OperationCanceledException = Exception
    ISelectionFilter = object
    ObjectType = None

from . import config, logger, task_models


MARKER_FAMILY_NAME = "CPSK_Маркер задания"
TEXT_LIMIT = 950
WRITABLE_MARKER_PARAMS = [
    u"CPSK_Дата",
    u"CPSK_Статус",
    u"CPSK_Пометка",
    u"CPSK_Комментарии",
    u"CPSK_Дисциплина",
    u"Новое",
    u"Принято",
    u"Отменено",
]


class CpskMarkerSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return is_cpsk_marker(element)

    def AllowReference(self, reference, position):
        return False


def is_cpsk_marker(element):
    try:
        if not element or not element.Category:
            return False
        if element.Category.Id.IntegerValue != int(BuiltInCategory.OST_DetailComponents):
            return False
        return marker_family_name(element) == MARKER_FAMILY_NAME
    except Exception:
        return False


def marker_family_name(element):
    try:
        return element.Symbol.Family.Name
    except Exception:
        return ""


def pick_marker(uidoc):
    message = u"Выберите на активном виде маркер задания.\nКатегория: Элементы узлов.\nСемейство: CPSK_Маркер задания."
    try:
        ref = uidoc.Selection.PickObject(ObjectType.Element, CpskMarkerSelectionFilter(), message)
        marker = uidoc.Document.GetElement(ref.ElementId)
        if not is_cpsk_marker(marker):
            return None, u"Выбран не маркер CPSK_Маркер задания"
        return marker, ""
    except OperationCanceledException:
        return None, u""
    except Exception:
        logger.exception("Marker selection failed")
        return None, u"Маркер не выбран"


def marker_datetime_text():
    return datetime.now().strftime("%d.%m.%Y %H:%M")


def set_marker_parameter_safe(element, param_name, value):
    try:
        param = element.LookupParameter(param_name)
        if not param:
            logger.write("Marker parameter not found: %s" % param_name)
            return False
        if param.IsReadOnly:
            logger.write("Marker parameter is read-only: %s" % param_name)
            return False
        storage = param.StorageType
        if storage == StorageType.Integer:
            param.Set(1 if value in (True, 1, "1", "true", "True", u"Да") else 0)
        elif storage == StorageType.Double:
            logger.write("Skip numeric marker parameter: %s" % param_name)
            return False
        else:
            param.Set("" if value is None else unicode(value))
        return True
    except Exception:
        logger.exception("Failed to set marker parameter: %s" % param_name)
        return False


def marker_parameter_text(element, param_name):
    try:
        param = element.LookupParameter(param_name)
        if not param:
            return u""
        if param.StorageType == StorageType.Integer:
            return unicode(param.AsInteger())
        if param.StorageType == StorageType.Double:
            try:
                return unicode(param.AsValueString() or "")
            except Exception:
                return unicode(param.AsDouble())
        try:
            return unicode(param.AsString() or param.AsValueString() or "")
        except Exception:
            return unicode(param.AsValueString() or "")
    except Exception:
        return u""


def marker_parameter_values(element):
    values = {}
    for param_name in WRITABLE_MARKER_PARAMS:
        value = marker_parameter_text(element, param_name)
        if value not in (None, u""):
            values[param_name] = value
    return values


def marker_status(status):
    return task_models.server_to_marker_status(status)


def apply_marker_status(element, status):
    status = marker_status(status)
    set_marker_parameter_safe(element, u"CPSK_Статус", status)
    set_marker_parameter_safe(element, u"Новое", status == u"Новое")
    set_marker_parameter_safe(element, u"Принято", status == u"Принято")
    set_marker_parameter_safe(element, u"Отменено", status == u"Отменено")


def _comment_text(task):
    parts = []
    if task.get("description"):
        parts.append(task.get("description"))
    if task.get("comment"):
        parts.append(task.get("comment"))
    text = "\n".join(parts)
    if len(text) > TEXT_LIMIT:
        logger.write("Marker comments truncated for task %s" % (task.get("id") or task.get("temporary_id") or ""))
        return text[:TEXT_LIMIT - 1]
    return text


def write_task_to_marker(doc, marker, task, update_date=True):
    if not marker:
        raise ValueError(u"Маркер не выбран")
    if not is_cpsk_marker(marker):
        raise ValueError(u"Выбран не маркер CPSK_Маркер задания")
    tx = Transaction(doc, "Записать задание в маркер")
    tx.Start()
    try:
        apply_marker_status(marker, task.get("marker_status") or task.get("status") or u"Новое")
        if update_date:
            set_marker_parameter_safe(marker, u"CPSK_Дата", marker_datetime_text())
        set_marker_parameter_safe(marker, u"CPSK_Пометка", task.get("type") or u"прочее")
        set_marker_parameter_safe(marker, u"CPSK_Комментарии", _comment_text(task))
        set_marker_parameter_safe(marker, u"CPSK_Дисциплина", u"%s → %s" % (task.get("sender_discipline") or "", task.get("receiver_discipline") or ""))
        tx.Commit()
    except Exception:
        tx.RollBack()
        logger.exception("Failed to write task to marker")
        raise


def find_marker_by_unique_id(doc, unique_id):
    if not unique_id:
        return None
    try:
        return doc.GetElement(unique_id)
    except Exception:
        return None
