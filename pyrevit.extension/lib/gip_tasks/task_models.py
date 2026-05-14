# -*- coding: utf-8 -*-
import uuid
from datetime import datetime


TASK_TYPES = [u"отверстие", u"проём", u"закладная", u"фундамент", u"площадка", u"коллизия", u"уточнение", u"замечание", u"прочее"]
MARKER_STATUSES = [u"Новое", u"Принято", u"Отменено"]
PRIORITIES = [u"низкий", u"обычный", u"высокий", u"срочный"]
DISCIPLINES = [u"КЖ", u"КМ", u"АР", u"ОВ", u"ВК", u"ЭОМ", u"СС", u"ТХ"]


def now_iso():
    return datetime.utcnow().replace(microsecond=0).isoformat() + "Z"


def now_local_text():
    return datetime.now().strftime("%d.%m.%Y %H:%M")


def server_to_marker_status(status):
    if status in ("cancelled", "returned", u"Отменено"):
        return u"Отменено"
    if status in ("in_progress", "review", "closed", u"Принято"):
        return u"Принято"
    return u"Новое"


def marker_to_server_status(status):
    if status == u"Принято":
        return "in_progress"
    if status == u"Отменено":
        return "cancelled"
    return "new"


def new_project_payload(settings, form_data):
    return {
        "id": form_data.get("id") or "",
        "company_id": settings.get("COMPANY_ID"),
        "name": form_data.get("name") or "",
        "code": form_data.get("code") or "",
        "customer": form_data.get("customer") or "",
        "address": form_data.get("address") or "",
        "description": form_data.get("description") or "",
        "created_at": form_data.get("created_at") or now_iso(),
        "updated_at": now_iso(),
        "is_local": bool(form_data.get("is_local", False)),
        "is_active": bool(form_data.get("is_active", True)),
        "is_archived": False,
    }


def new_task_payload(settings, form_data, revit_context):
    project_id = settings.get("CURRENT_PROJECT_ID") or form_data.get("project_id") or ""
    if not project_id:
        raise ValueError(u"Сначала выберите проект")
    marker = (revit_context or {}).get("marker") or {}
    view = (revit_context or {}).get("active_view") or {}
    title = form_data.get("type") or u"Задание"
    marker_status = form_data.get("marker_status") or u"Новое"
    payload = {
        "id": form_data.get("id") or "",
        "company_id": settings.get("COMPANY_ID"),
        "project_id": project_id,
        "project_name": settings.get("CURRENT_PROJECT_NAME"),
        "project_code": settings.get("CURRENT_PROJECT_CODE"),
        "title": title,
        "description": form_data.get("description") or "",
        "comment": form_data.get("comment") or "",
        "type": form_data.get("type") or u"прочее",
        "status": marker_to_server_status(marker_status),
        "marker_status": marker_status,
        "priority": form_data.get("priority") or u"обычный",
        "sender_user_id": settings.get("CURRENT_USER"),
        "sender_discipline": form_data.get("sender_discipline") or settings.get("CURRENT_DISCIPLINE"),
        "receiver_discipline": form_data.get("receiver_discipline") or "",
        "due_date": form_data.get("due_date") or "",
        "created_at": form_data.get("created_at") or now_iso(),
        "updated_at": now_iso(),
        "source_model_name": (revit_context or {}).get("document_title") or "",
        "current_revit_file_path": (revit_context or {}).get("document_path") or "",
        "marker_element_id": marker.get("element_id") or "",
        "marker_unique_id": marker.get("unique_id") or "",
        "marker_family_name": marker.get("family_name") or "",
        "marker_coordinates": marker.get("location") or {},
        "level_name": view.get("level") or "",
        "view_name": view.get("name") or "",
        "view_id": view.get("id") or "",
        "is_deleted": False,
    }
    if not payload["id"]:
        payload["temporary_id"] = "local-task-" + str(uuid.uuid4())
    return payload

